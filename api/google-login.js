// 구글 로그인 콜백 — Google Identity Services가 credential(ID 토큰)과 g_csrf_token을 POST로 보내온다.
const { sessionCookie, verifyGoogleLogin } = require('../lib/web');

module.exports = async (req, res) => {
  const body = req.body || {};
  const csrf = ((req.headers.cookie || '').match(/g_csrf_token=([^;]+)/) || [])[1];
  // 이동할 경로는 같은 사이트 안(/로 시작, //로 시작하지 않음)만 허용
  const next = /^\/(?!\/)/.test(String((req.query || {}).next || '')) ? req.query.next.replace(/^\/api\/index/, '/') : '/';
  const fail = msg => res.redirect(303, '/?login_error=' + encodeURIComponent(msg));
  if (req.method !== 'POST' || !csrf || csrf !== body.g_csrf_token) return fail('로그인 요청이 올바르지 않습니다. 다시 시도하세요.');
  const r = await verifyGoogleLogin(body.credential);
  if (r.error) return fail(r.error);
  res.setHeader('Set-Cookie', sessionCookie(r.email));
  res.redirect(303, next);
};
