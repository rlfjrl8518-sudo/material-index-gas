// 웹앱 진입점(doGet 대체) + 비밀번호 로그인
const fs = require('fs'), path = require('path');
const { accessToken, sessionCookie, isAuthed, checkPassword, appUrl, runGas } = require('../lib/web');

const INDEX = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');

// 브라우저용 google.script.run 대체 — withSuccessHandler/withFailureHandler 체이닝을 그대로 지원한다.
const POLYFILL = `<script>
window.google = { script: { run: (function mk(ok, fail) {
  return new Proxy({}, { get: function (_, k) {
    if (k === 'withSuccessHandler') return function (h) { return mk(h, fail); };
    if (k === 'withFailureHandler') return function (h) { return mk(ok, h); };
    if (k === 'withUserObject') return function () { return mk(ok, fail); };
    return function () {
      var args = Array.prototype.slice.call(arguments);
      fetch('/api/rpc', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: k, args: args }) })
        .then(function (r) { return r.json(); })
        .then(function (j) { if (j.error !== undefined) { if (fail) fail(new Error(j.error)); } else if (ok) ok(j.result); },
              function (e) { if (fail) fail(e); });
    };
  } });
})() } };
</script>`;

// 로그인 화면 — GOOGLE_LOGIN_CLIENT_ID가 있으면 구글 로그인 버튼을, APP_PASSWORD가 있으면 비밀번호 칸을 보여준다.
const LOGIN = (msg, next) => {
  const gid = process.env.GOOGLE_LOGIN_CLIENT_ID;
  const google = gid ? `<script src="https://accounts.google.com/gsi/client" async></script>
<div id="g_id_onload" data-client_id="${gid}" data-login_uri="/api/google-login${next ? '?next=' + encodeURIComponent(next) : ''}" data-ux_mode="redirect" data-auto_prompt="false"></div>
<div class="g_id_signin" data-type="standard" data-size="large" data-width="256" data-text="signin_with" data-locale="ko"></div>` : '';
  const pw = process.env.APP_PASSWORD ? `${gid ? '<div class="or">또는</div>' : ''}<input type="password" name="password" placeholder="비밀번호" ${gid ? '' : 'autofocus'} required>
<button>로그인</button>` : '';
  const esc = t => String(t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>운영 소재 분석 — 로그인</title>
<style>body{font-family:system-ui,sans-serif;background:#f5f6f8;display:grid;place-items:center;min-height:100vh;margin:0}
form{background:#fff;padding:32px;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,.08);width:min(320px,90vw)}
h1{font-size:18px;margin:0 0 16px}input{width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccd;border-radius:8px;font-size:15px}
button{margin-top:12px;width:100%;padding:10px;border:0;border-radius:8px;background:#2563eb;color:#fff;font-size:15px;cursor:pointer}
.or{text-align:center;color:#9ca3af;font-size:12px;margin:14px 0}p{color:#c0392b;font-size:13px;margin:8px 0 0}</style></head><body>
<form method="post"><h1>운영 소재 분석</h1>${google}${pw}${msg ? '<p>' + esc(msg) + '</p>' : ''}</form></body></html>`;
};

const html = (res, code, body) => res.status(code).setHeader('Content-Type', 'text/html; charset=utf-8').setHeader('Cache-Control', 'no-store').send(body);

module.exports = async (req, res) => {
  if (req.method === 'POST') {
    if (!checkPassword((req.body || {}).password)) return html(res, 401, LOGIN('비밀번호가 맞지 않습니다.', req.url));
    res.setHeader('Set-Cookie', sessionCookie('비밀번호 로그인'));
    const url = req.url || '';
    return res.redirect(303, '/' + (url.includes('?') ? url.slice(url.indexOf('?')) : ''));
  }
  const user = isAuthed(req);
  if (!user) return html(res, 200, LOGIN(String((req.query || {}).login_error || '').slice(0, 200), req.url));

  // doGet과 동일: ?unit= 이 있으면 번들 데이터를 미리 실어 보낸다.
  const q = req.query || {};
  let bundleData = null;
  if (q.unit) {
    bundleData = (await runGas({ fn: 'getBundleData', args: [q.unit, { 매체: q._media || '', 캠페인: q._campaign || '', 그룹: q._group || '', 소재이름: q._name || '' }],
      token: await accessToken(), appUrl: appUrl(req), user })).result;
  }
  const page = INDEX
    .replace('<?!= JSON.stringify(bundleData || null) ?>', () => JSON.stringify(bundleData).replace(/</g, '\\u003c'))
    .replace(/<head>/i, () => '<head><title>' + (q.unit ? '광고단위 소재' : '운영 소재 분석') + '</title>' + POLYFILL);
  html(res, 200, page);
};
