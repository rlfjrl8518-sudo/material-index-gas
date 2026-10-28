// Vercel 함수 공통: Google 액세스 토큰(refresh token으로 발급), 비밀번호 세션 쿠키, 응답 헬퍼
const crypto = require('crypto');

const COOKIE = 'mi_session';
const SESSION_DAYS = 30;
const tokenCache = {};

// scope를 주면 그 범위로 축소한 토큰을 받는다 — 브라우저에 넘기는 업로드용 토큰은 drive로만 제한.
async function accessToken(scope) {
  const key = scope || 'full';
  const c = tokenCache[key];
  if (c && c.exp > Date.now() + 60e3) return c.token;
  const body = new URLSearchParams({
    client_id: process.env.GOOGLE_CLIENT_ID,
    client_secret: process.env.GOOGLE_CLIENT_SECRET,
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
    grant_type: 'refresh_token'
  });
  if (scope) body.set('scope', scope);
  const res = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', body });
  const j = await res.json();
  if (!j.access_token) throw new Error('Google 토큰 발급 실패: ' + (j.error_description || j.error || res.status));
  tokenCache[key] = { token: j.access_token, exp: Date.now() + j.expires_in * 1000 };
  return j.access_token;
}

const secret = () => process.env.APP_SESSION_SECRET || '';
const sign = v => crypto.createHmac('sha256', secret()).update(v).digest('base64url');
const safeEqual = (a, b) => a.length === b.length && crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));

function sessionCookie() {
  const exp = String(Date.now() + SESSION_DAYS * 864e5);
  return `${COOKIE}=${exp}.${sign(exp)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${SESSION_DAYS * 86400}`;
}

function isAuthed(req) {
  if (!secret()) return false;
  const m = (req.headers.cookie || '').match(new RegExp(COOKIE + '=([0-9]+)\\.([\\w-]+)'));
  return !!m && Number(m[1]) > Date.now() && safeEqual(m[2], sign(m[1]));
}

const isInternal = req => !!secret() && safeEqual(String(req.headers['x-trigger-key'] || ''), sign('trigger'));
const triggerKey = () => sign('trigger');

function checkPassword(pw) {
  const want = process.env.APP_PASSWORD || '';
  return !!want && safeEqual(sign(String(pw || '')), sign(want));
}

const appUrl = req => process.env.APP_URL || ('https://' + req.headers.host);

module.exports = { accessToken, sessionCookie, isAuthed, isInternal, triggerKey, checkPassword, appUrl };

// .gs 함수 하나를 워커에서 실행하고 { result, triggers }를 돌려준다.
function runGas(workerData) {
  const { Worker } = require('worker_threads');
  return new Promise((resolve, reject) => {
    const w = new Worker(require('path').join(__dirname, 'worker.js'), { workerData });
    w.once('message', m => {
      if (m.error) { console.error(workerData.fn, m.stack); reject(new Error(m.error)); }
      else resolve({ result: JSON.parse(m.result), triggers: m.triggers, timing: m.timing });
    });
    w.once('error', reject);
    w.once('exit', code => { if (code) reject(new Error('worker exit ' + code)); });
  });
}

module.exports.runGas = runGas;
