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

// 세션 쿠키: 만료시각.사용자(base64url).서명 — 사용자는 구글 로그인 이메일 또는 '비밀번호 로그인'
function sessionCookie(user) {
  const exp = String(Date.now() + SESSION_DAYS * 864e5);
  const u = Buffer.from(String(user || '비밀번호 로그인')).toString('base64url');
  return `${COOKIE}=${exp}.${u}.${sign(exp + '.' + u)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${SESSION_DAYS * 86400}`;
}

// 로그인돼 있으면 사용자 이름(이메일)을, 아니면 null을 돌려준다.
function isAuthed(req) {
  if (!secret()) return null;
  const m = (req.headers.cookie || '').match(new RegExp(COOKIE + '=([0-9]+)\\.([\\w-]*)\\.([\\w-]+)'));
  if (!m || Number(m[1]) <= Date.now() || !safeEqual(m[3], sign(m[1] + '.' + m[2]))) return null;
  return Buffer.from(m[2], 'base64url').toString() || '비밀번호 로그인';
}

// 구글 로그인(ID 토큰) 검증 — 허용 목록: ALLOWED_EMAILS(쉼표 구분) 또는 ALLOWED_DOMAINS(예: playd.com)
async function verifyGoogleLogin(idToken) {
  const clientId = process.env.GOOGLE_LOGIN_CLIENT_ID;
  if (!clientId || !idToken) return { error: '구글 로그인이 설정되지 않았습니다.' };
  const r = await fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken));
  const t = await r.json();
  if (!r.ok || t.aud !== clientId || t.email_verified !== 'true' || Number(t.exp) * 1000 < Date.now()) return { error: '구글 로그인 확인에 실패했습니다.' };
  const email = String(t.email).toLowerCase();
  const list = v => String(process.env[v] || '').toLowerCase().split(',').map(x => x.trim()).filter(Boolean);
  if (!list('ALLOWED_EMAILS').includes(email) && !list('ALLOWED_DOMAINS').includes(email.split('@')[1])) return { error: email + ' 계정은 접근 권한이 없습니다. 관리자에게 요청하세요.' };
  return { email };
}

const isInternal = req => !!secret() && safeEqual(String(req.headers['x-trigger-key'] || ''), sign('trigger'));
const triggerKey = () => sign('trigger');

function checkPassword(pw) {
  const want = process.env.APP_PASSWORD || '';
  return !!want && safeEqual(sign(String(pw || '')), sign(want));
}

const appUrl = req => process.env.APP_URL || ('https://' + req.headers.host);

module.exports = { accessToken, sessionCookie, isAuthed, isInternal, triggerKey, checkPassword, appUrl, verifyGoogleLogin };

// .gs 함수 하나를 워커 스레드에서 실행하고 { result, triggers, timing }을 돌려준다.
// 스레드는 재사용한다(요청마다 새로 만들면 시트 캐시가 매번 비어 분당 읽기 한도에 금방 닿는다).
// 한 스레드는 한 번에 한 요청만 처리하고, 모두 바쁘면 POOL_MAX까지 늘린다.
const POOL_MAX = 4;
const pool = [];

function spawnWorker() {
  const { Worker } = require('worker_threads');
  const w = { worker: new Worker(require('path').join(__dirname, 'worker.js')), current: null, queue: [] };
  const fail = e => {
    pool.splice(pool.indexOf(w), 1);
    [w.current].concat(w.queue).filter(Boolean).forEach(t => t.reject(e));
  };
  w.worker.on('error', fail);
  w.worker.on('exit', code => { if (pool.indexOf(w) !== -1) fail(new Error('worker exit ' + code)); });
  w.worker.on('message', m => {
    const t = w.current;
    w.current = null;
    if (m.error) { console.error(t.job.fn, m.stack); t.reject(new Error(m.error)); }
    else t.resolve({ result: JSON.parse(m.result), triggers: m.triggers, timing: m.timing });
    startNext(w);
  });
  pool.push(w);
  return w;
}

function startNext(w) {
  if (w.current || !w.queue.length) return;
  w.current = w.queue.shift();
  w.worker.postMessage(w.current.job);
}

function runGas(job) {
  const w = pool.find(x => !x.current && !x.queue.length)
    || (pool.length < POOL_MAX ? spawnWorker() : pool.reduce((a, b) => (a.queue.length <= b.queue.length ? a : b)));
  return new Promise((resolve, reject) => {
    w.queue.push({ job, resolve, reject });
    startNext(w);
  });
}

module.exports.runGas = runGas;
