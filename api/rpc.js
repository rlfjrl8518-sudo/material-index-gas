// google.script.run 대체 — POST {fn, args} → 기존 .gs 함수를 그대로 실행해 {result} 또는 {error}로 돌려준다.
const zlib = require('zlib');
const { waitUntil } = require('@vercel/functions');
const { accessToken, isAuthed, isInternal, triggerKey, appUrl, runGas } = require('../lib/web');

// 브라우저가 Drive로 직접 업로드할 때 쓰는 토큰을 돌려주는 함수들 — 전체 권한 토큰 대신 drive 범위로 축소해서 넘긴다.
const UPLOAD_TOKEN_FNS = ['getUploadToken', 'getInspectionUploadContext'];
// 시트 메뉴 전용(getUi 사용) 또는 웹앱 진입점 — 대시보드에서 호출할 이유가 없다.
const BLOCKED = ['onOpen', 'doGet', 'openDashboard'];
// Vercel 함수 응답 한도(4.5MB)는 압축 전 크기에 걸리므로, 큰 응답은 직접 gzip해서 보낸다(소재_통합RAW 3MB → 약 0.3MB).
const GZIP_OVER = 256 * 1024;

async function run(fn, args, req, user, res) {
  const t0 = Date.now();
  const token = await accessToken();
  const t1 = Date.now();
  const { result, triggers, timing } = await runGas({ fn, args, token, appUrl: appUrl(req), user, audit: true });
  if (res) res.setHeader('Server-Timing', `token;dur=${t1 - t0}, worker;dur=${Date.now() - t1}, gas;dur=${timing.gas}, http;dur=${timing.http};desc="${timing.calls} calls"`);
  if (result && result.token && UPLOAD_TOKEN_FNS.includes(fn)) result.token = await accessToken('https://www.googleapis.com/auth/drive');
  // ScriptApp.newTrigger(...).after(ms) 대체: 응답 후 같은 함수를 자기 자신에게 다시 호출한다.
  triggers.forEach(t => waitUntil(new Promise(r => setTimeout(r, t.ms)).then(() => fetch(appUrl(req) + '/api/rpc', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-trigger-key': triggerKey() },
    body: JSON.stringify({ fn: t.fn, args: [] })
  })).catch(e => console.error('trigger 실패', t.fn, e))));
  return result;
}

function sendJson(req, res, obj) {
  const body = JSON.stringify(obj);
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  if (body.length > GZIP_OVER && /\bgzip\b/.test(req.headers['accept-encoding'] || '')) {
    res.setHeader('Content-Encoding', 'gzip');
    return res.status(200).end(zlib.gzipSync(body));
  }
  res.status(200).end(body);
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).end();
  const internal = isInternal(req);
  const user = internal ? '자동 처리' : isAuthed(req);
  if (!user) return res.status(401).json({ error: '로그인이 필요합니다. 페이지를 새로고침하세요.' });
  const { fn, args } = req.body || {};
  if (typeof fn !== 'string' || fn.startsWith('_') || BLOCKED.includes(fn)) return res.status(400).json({ error: '허용되지 않은 함수: ' + fn });

  // 트리거 호출은 바로 202로 답하고 실제 처리는 waitUntil에서 계속한다(호출한 쪽이 기다리지 않게).
  if (internal) {
    waitUntil(run(fn, args, req, user).catch(e => console.error('trigger 처리 실패', fn, e)));
    return res.status(202).json({ ok: true });
  }
  try {
    sendJson(req, res, { result: (await run(fn, args, req, user, res)) ?? null });
  } catch (e) {
    console.error(fn, e);
    sendJson(req, res, { error: e.message || String(e) });
  }
};
