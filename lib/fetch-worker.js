// 동기 HTTP용 상주 스레드 — gas.js가 요청을 보내고 Atomics.wait로 멈춰 있는 동안 여기서 fetch를 돌린다.
// 매 호출마다 프로세스를 띄우던 방식보다 기동 비용이 없고, 같은 호스트로의 연결(keep-alive)을 재사용한다.
const { workerData } = require('worker_threads');
const { port, flag } = workerData;
const FETCH_TIMEOUT_MS = 150000; // 응답 없는 요청 하나가 함수 전체(300초)를 붙잡지 않게

port.on('message', async reqs => {
  const out = await Promise.all(reqs.map(async r => {
    try {
      const res = await fetch(r.url, {
        method: r.method || 'GET',
        headers: r.headers || {},
        body: r.body64 ? Buffer.from(r.body64, 'base64') : r.body,
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS)
      });
      return { code: res.status, type: res.headers.get('content-type') || '', body: new Uint8Array(await res.arrayBuffer()) };
    } catch (e) {
      return { code: 0, type: '', body: new TextEncoder().encode(String((e && e.cause) || e)) };
    }
  }));
  port.postMessage(out);
  Atomics.store(flag, 0, 1);
  Atomics.notify(flag, 0);
});
