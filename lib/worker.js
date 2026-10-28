// GAS 실행 스레드 — 요청을 하나씩 받아 처리한다(스레드를 재사용해 시트 캐시·HTTP 연결·컴파일 결과를 유지).
// GAS 코드의 동기 HTTP(Atomics.wait)가 메인 이벤트 루프를 막지 않도록 메인 스레드 밖에서 돈다.
const { parentPort } = require('worker_threads');
const { createRuntime, httpStats } = require('./gas');

parentPort.on('message', job => {
  const t0 = Date.now();
  httpStats.ms = 0; httpStats.n = 0;
  try {
    const rt = createRuntime({ token: job.token, appUrl: job.appUrl, user: job.user, audit: job.audit });
    if (typeof rt.ctx[job.fn] !== 'function') throw new Error('알 수 없는 함수: ' + job.fn);
    const result = rt.call(job.fn, job.args);
    // vm 컨텍스트의 객체를 그대로 넘기면 structured clone이 일부 값을 바꾸므로 JSON으로 고정한다(google.script.run도 직렬화한다).
    const json = JSON.stringify(result === undefined ? null : result);
    parentPort.postMessage({ result: json, triggers: rt.triggers, timing: { gas: Date.now() - t0, http: httpStats.ms, calls: httpStats.n } });
  } catch (e) {
    parentPort.postMessage({ error: (e && e.message) || String(e), stack: e && e.stack });
  }
});
