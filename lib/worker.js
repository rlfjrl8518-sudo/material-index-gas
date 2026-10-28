// 요청 하나를 워커 스레드에서 실행한다 — GAS 코드의 동기 HTTP(execFileSync)가 메인 이벤트 루프를 막지 않도록.
const { parentPort, workerData } = require('worker_threads');
const { createRuntime } = require('./gas');

const { fn, args, token, appUrl, user } = workerData;
try {
  const rt = createRuntime({ token, appUrl, user });
  if (typeof rt.ctx[fn] !== 'function') throw new Error('알 수 없는 함수: ' + fn);
  const result = rt.call(fn, args);
  // vm 컨텍스트의 객체를 그대로 넘기면 structured clone이 일부 값을 바꾸므로 JSON으로 고정한다(google.script.run도 직렬화한다).
  parentPort.postMessage({ result: JSON.stringify(result === undefined ? null : result), triggers: rt.triggers });
} catch (e) {
  parentPort.postMessage({ error: (e && e.message) || String(e), stack: e && e.stack });
}
