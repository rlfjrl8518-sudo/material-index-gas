// stdin: [{url, method, headers, body, body64}] → stdout: [{code, type, body64}]
// gas.js가 동기 HTTP를 흉내 내려고 이 파일을 자식 프로세스로 실행하고 끝날 때까지 기다린다.
let input = '';
process.stdin.on('data', d => input += d).on('end', async () => {
  const out = await Promise.all(JSON.parse(input).map(async r => {
    try {
      const res = await fetch(r.url, {
        method: r.method || 'GET',
        headers: r.headers || {},
        body: r.body64 ? Buffer.from(r.body64, 'base64') : r.body
      });
      return { code: res.status, type: res.headers.get('content-type') || '', body64: Buffer.from(await res.arrayBuffer()).toString('base64') };
    } catch (e) {
      return { code: 0, type: '', body64: Buffer.from(String(e && e.cause || e)).toString('base64') };
    }
  }));
  process.stdout.write(JSON.stringify(out));
});
