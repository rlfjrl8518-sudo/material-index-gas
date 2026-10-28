// GAS 런타임 호환 레이어 — 기존 .gs 파일을 고치지 않고 Node(vm)에서 그대로 실행한다.
// GAS API는 전부 동기라서 HTTP는 상주 스레드(fetch-worker.js)에 맡기고 Atomics.wait로 기다리는 방식으로 흉내 낸다.
// 시트는 시트별로 처음 접근할 때 한 번 읽어 메모리 grid에 두고, 쓰기는 순서대로 모았다가 요청 끝에 일괄 반영한다.
process.env.TZ = 'Asia/Seoul';
const fs = require('fs'), path = require('path'), vm = require('vm');
const { unzipSync } = require('fflate');

const ROOT = path.join(__dirname, '..');
const TZ = 'Asia/Seoul';
const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets/';
const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';
const PROPS_SHEET = '_properties';
const AUDIT_SHEET = '작업이력';
const GS_FILES = fs.readdirSync(ROOT).filter(f => f.endsWith('.gs')).sort();
const scripts = GS_FILES.map(f => new vm.Script(fs.readFileSync(path.join(ROOT, f), 'utf8'), { filename: f }));

// ---------- HTTP ----------
// GAS는 동기 API라서, 별도 스레드(fetch-worker.js)에 요청을 넘기고 응답이 올 때까지 Atomics.wait로 기다린다.
let fetcher = null;
const httpStats = { ms: 0, n: 0 }; // Server-Timing 헤더용
function httpAll(reqs) {
  if (!fetcher) {
    const { Worker, MessageChannel } = require('worker_threads');
    const { port1, port2 } = new MessageChannel();
    const flag = new Int32Array(new SharedArrayBuffer(4));
    const worker = new Worker(path.join(__dirname, 'fetch-worker.js'), { workerData: { port: port2, flag }, transferList: [port2] });
    worker.unref();
    port1.unref();
    fetcher = { port: port1, flag };
  }
  const { receiveMessageOnPort } = require('worker_threads');
  const t = Date.now();
  if (process.env.GAS_DEBUG) console.log('[http]', reqs.length, reqs.map(r => r.method + ' ' + r.url.split('?')[0].slice(-60)).join(' | '));
  Atomics.store(fetcher.flag, 0, 0);
  fetcher.port.postMessage(reqs);
  if (Atomics.wait(fetcher.flag, 0, 0, 170000) === 'timed-out') {
    fetcher = null; // 멈춘 스레드는 버리고 다음 호출 때 새로 만든다
    throw new Error('외부 호출 응답 없음(170초): ' + reqs.map(r => r.url.split('?')[0]).slice(0, 3).join(', '));
  }
  httpStats.ms += Date.now() - t; httpStats.n++;
  return receiveMessageOnPort(fetcher.port).message.map(r => ({ code: r.code, type: r.type, body: Buffer.from(r.body.buffer, r.body.byteOffset, r.body.byteLength) }));
}

// Sheets/Drive는 분당 한도(사용자당 읽기 60회 등)가 있어 429가 나면 잠깐 기다렸다 다시 시도한다.
function httpRetry(reqs) {
  let out = httpAll(reqs);
  for (let wait = 5000; wait <= 20000 && out.some(r => r.code === 429); wait *= 2) {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, wait);
    const again = httpAll(reqs.filter((_, i) => out[i].code === 429));
    let k = 0;
    out = out.map(r => (r.code === 429 ? again[k++] : r));
  }
  return out;
}

function google(token, method, url, json) {
  const req = { method, url, headers: { Authorization: 'Bearer ' + token } };
  if (json !== undefined) { req.headers['Content-Type'] = 'application/json'; req.body = JSON.stringify(json); }
  const r = httpRetry([req])[0];
  const text = r.body.toString('utf8');
  if (r.code < 200 || r.code >= 300) {
    const err = new Error(method + ' ' + url.split('?')[0] + ' → ' + r.code + ' ' + text.slice(0, 300));
    err.code = r.code;
    throw err;
  }
  return text ? JSON.parse(text) : {};
}

// ---------- Blob / Utilities ----------
class Blob {
  constructor(bytes, type, name) { this.bytes = Buffer.from(bytes || []); this.type = type || 'application/octet-stream'; this.name = name || ''; }
  getBytes() { return this.bytes; }
  getContentType() { return this.type; }
  setContentType(t) { this.type = t; return this; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getDataAsString() { return this.bytes.toString('utf8'); }
  getBlob() { return this; }
}

const pad = (n, w = 2) => String(n).padStart(w, '0');
// Intl.DateTimeFormat 생성이 비싸서(소재_통합RAW 7천여 행 날짜 변환에 1초) 타임존별로 하나만 만든다.
const dtfCache = {};
function formatDate(date, tz, fmt) {
  const p = {};
  tz = tz || TZ;
  (dtfCache[tz] || (dtfCache[tz] = new Intl.DateTimeFormat('en-US', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })))
    .formatToParts(new Date(date)).forEach(x => { p[x.type] = x.value; });
  const map = { yyyy: p.year, yy: p.year.slice(2), MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second };
  return fmt.replace(/yyyy|yy|MM|dd|HH|mm|ss/g, t => map[t]);
}

const Utilities = {
  formatDate,
  base64Encode: d => Buffer.from(d).toString('base64'),
  base64Decode: s => Buffer.from(s, 'base64'),
  newBlob: (data, type, name) => new Blob(typeof data === 'string' ? Buffer.from(data, 'utf8') : data, type, name),
  sleep: ms => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms),
  DigestAlgorithm: { SHA_256: 'sha256', MD5: 'md5', SHA_1: 'sha1' },
  // GAS처럼 부호 있는 바이트 배열(-128~127)을 돌려준다.
  computeDigest: (alg, data) => Array.from(require('crypto').createHash(alg).update(Buffer.from(data)).digest()).map(b => (b > 127 ? b - 256 : b)),
  unzip: blob => Object.entries(unzipSync(new Uint8Array(blob.getBytes())))
    .filter(([n]) => !n.endsWith('/'))
    .map(([n, b]) => new Blob(Buffer.from(b), '', n))
};

// ---------- UrlFetchApp ----------
function toHttpReq(url, o) {
  o = o || {};
  const headers = Object.assign({}, o.headers);
  let body = o.payload;
  if (body != null && typeof body === 'object' && !Buffer.isBuffer(body)) {
    body = new URLSearchParams(body).toString();
    headers['Content-Type'] = headers['Content-Type'] || 'application/x-www-form-urlencoded';
  } else if (o.contentType) headers['Content-Type'] = o.contentType;
  return { url, method: (o.method || (body != null ? 'post' : 'get')).toUpperCase(), headers, body };
}
function toHttpResponse(r, req, mute) {
  if (!mute && (r.code < 200 || r.code >= 400)) throw new Error('Request failed for ' + req.url + ' returned code ' + r.code + '. ' + r.body.toString('utf8').slice(0, 300));
  return {
    getResponseCode: () => r.code,
    getContentText: () => r.body.toString('utf8'),
    getBlob: () => new Blob(r.body, (r.type || '').split(';')[0]),
    getHeaders: () => ({ 'Content-Type': r.type })
  };
}
const UrlFetchApp = {
  fetch: (url, o) => { const req = toHttpReq(url, o); return toHttpResponse(httpAll([req])[0], req, o && o.muteHttpExceptions); },
  fetchAll: list => {
    const reqs = list.map(x => (typeof x === 'string' ? toHttpReq(x) : toHttpReq(x.url, x)));
    return httpAll(reqs).map((r, i) => toHttpResponse(r, reqs[i], list[i].muteHttpExceptions));
  }
};

// ---------- SpreadsheetApp ----------
const SERIAL_EPOCH = Date.UTC(1899, 11, 30);
const KST_OFFSET = 9 * 3600e3;
// 날짜 칸 판별: 같은 범위를 SERIAL_NUMBER / FORMATTED_STRING 두 방식으로 받아,
// 앞은 숫자인데 뒤는 문자열인 칸이 날짜다(dateTimeRenderOption은 날짜·시간 서식 칸에만 적용됨).
// 칸별 서식까지 받는 spreadsheets.get은 소재_통합RAW 기준 25MB/13초라 쓰지 않는다(values는 gzip 0.4MB/1초).
// D: .gs 코드가 instanceof Date로 검사하므로 vm 컨텍스트의 Date로 만들어야 한다(realm이 다르면 instanceof가 false).
function cellValue(serial, text, D) {
  if (serial === undefined || serial === null) return '';
  if (typeof serial === 'number' && typeof text === 'string') return new D(SERIAL_EPOCH + serial * 864e5 - KST_OFFSET);
  return serial;
}
const isDate = v => Object.prototype.toString.call(v) === '[object Date]';
const writeValue = v => (isDate(v) ? formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss') : v == null ? '' : v);
const isEmpty = v => v === '' || v == null;
function colA1(n) { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; }

class Range {
  constructor(sheet, row, col, nr, nc) { Object.assign(this, { sheet, row, col, nr: nr || 1, nc: nc || 1 }); }
  getValues() {
    const g = this.sheet._grid();
    return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => {
      const r = g[this.row - 1 + i];
      const v = r ? r[this.col - 1 + j] : undefined;
      return v === undefined ? '' : v;
    }));
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(vals) {
    if (vals.length !== this.nr || vals.some(r => r.length !== this.nc)) {
      throw new Error('데이터의 행/열 수(' + vals.length + 'x' + ((vals[0] || []).length) + ')가 범위(' + this.nr + 'x' + this.nc + ')와 맞지 않습니다.');
    }
    const g = this.sheet._grid();
    vals.forEach((r, i) => {
      const gr = g[this.row - 1 + i] || (g[this.row - 1 + i] = []);
      r.forEach((v, j) => { gr[this.col - 1 + j] = v == null ? '' : v; });
    });
    for (let i = 0; i < g.length; i++) if (!g[i]) g[i] = [];
    this.sheet._ensureSize(this.row + this.nr - 1, this.col + this.nc - 1);
    this.sheet.ss._op({
      values: {
        range: "'" + this.sheet.name.replace(/'/g, "''") + "'!" + colA1(this.col) + this.row + ':' + colA1(this.col + this.nc - 1) + (this.row + this.nr - 1),
        values: vals.map(r => r.map(writeValue))
      }
    });
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  _gridRange() { return { sheetId: this.sheet.id, startRowIndex: this.row - 1, endRowIndex: this.row - 1 + this.nr, startColumnIndex: this.col - 1, endColumnIndex: this.col - 1 + this.nc }; }
  setFontWeight(w) {
    this.sheet.ss._op({ req: { repeatCell: { range: this._gridRange(), cell: { userEnteredFormat: { textFormat: { bold: w === 'bold' } } }, fields: 'userEnteredFormat.textFormat.bold' } } });
    return this;
  }
  insertCheckboxes() {
    this.sheet.ss._op({ req: { setDataValidation: { range: this._gridRange(), rule: { condition: { type: 'BOOLEAN' } } } } });
    return this;
  }
}

// 시트 읽기 캐시 — 같은 시트를 CACHE_TTL_MS 안에 다시 읽으면 API를 부르지 않는다(분당 읽기 한도 보호).
// 이 프로세스에서 쓰기를 반영(flush)하면 해당 시트 캐시는 바로 버린다. 시트를 직접 고친 내용은 최대 TTL만큼 늦게 보인다.
// 인스턴스·스레드마다 따로 들고 있는 캐시라 다른 곳에서 막 쓴 내용은 최대 TTL만큼 늦게 보인다. 그래서 대시보드가 '행 위치를 보고
// 쓰는'(appendRow 등) 시트는 캐시하지 않는다 — 오래된 마지막 행을 보고 덧붙이면 남이 방금 쓴 행을 덮어쓸 수 있다.
// 캐시 대상은 대시보드가 쓰지 않거나(RAW_*, 소재_통합RAW — 시트 메뉴의 통합 적재만 씀) 고정 칸만 쓰는(설정, 검수기준) 시트뿐이다.
const CACHE_TTL_MS = 20000;
const isCacheable = name => /^RAW_/.test(name) || ['소재_통합RAW', '설정', '검수기준'].indexOf(name) !== -1;
const sheetCache = new Map(); // ssId   시트명 → { at, serial, text } / ssId → { at, meta }
const cacheGet = k => { const c = sheetCache.get(k); return c && Date.now() - c.at < CACHE_TTL_MS ? c : null; };
function cacheDrop(ssId, sheetName) {
  if (sheetName === undefined) { for (const k of sheetCache.keys()) if (k === ssId || k.startsWith(ssId + ' ')) sheetCache.delete(k); }
  else sheetCache.delete(ssId + ' ' + sheetName);
}

class Sheet {
  constructor(ss, props, grid) {
    this.ss = ss; this.id = props.sheetId; this.name = props.title;
    this.rows = (props.gridProperties || {}).rowCount || 1000;
    this.cols = (props.gridProperties || {}).columnCount || 26;
    this.grid = grid || null;
  }
  _grid() {
    if (!this.grid) {
      const key = this.ss.id + ' ' + this.name;
      let c = isCacheable(this.name) ? cacheGet(key) : null;
      if (!c) {
        const url = SHEETS_API + this.ss.id + '/values/' + encodeURIComponent("'" + this.name.replace(/'/g, "''") + "'")
          + '?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=';
        const [serial, text] = httpRetry(['SERIAL_NUMBER', 'FORMATTED_STRING'].map(o => ({ method: 'GET', url: url + o, headers: { Authorization: 'Bearer ' + this.ss.token } })))
          .map(r => {
            if (r.code !== 200) throw new Error('시트 읽기 실패(' + this.name + ') → ' + r.code + ' ' + r.body.toString('utf8').slice(0, 300));
            return JSON.parse(r.body.toString('utf8')).values || [];
          });
        c = { at: Date.now(), serial, text };
        if (isCacheable(this.name)) sheetCache.set(key, c);
      }
      this.grid = c.serial.map((row, i) => row.map((v, j) => cellValue(v, (c.text[i] || [])[j], this.ss.Date)));
    }
    return this.grid;
  }
  _ensureSize(rows, cols) {
    if (rows > this.rows) { this.ss._op({ req: { appendDimension: { sheetId: this.id, dimension: 'ROWS', length: rows - this.rows + 100 } } }); this.rows = rows + 100; }
    if (cols > this.cols) { this.ss._op({ req: { appendDimension: { sheetId: this.id, dimension: 'COLUMNS', length: cols - this.cols } } }); this.cols = cols; }
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getLastRow() { const g = this._grid(); let i = g.length; while (i > 0 && !(g[i - 1] || []).some(v => !isEmpty(v))) i--; return i; }
  getLastColumn() { return this._grid().reduce((m, r) => { let j = r.length; while (j > 0 && isEmpty(r[j - 1])) j--; return Math.max(m, j); }, 0); }
  getMaxRows() { return this.rows; }
  getMaxColumns() { return this.cols; }
  getRange(row, col, nr, nc) {
    if (typeof row !== 'number') throw new Error('A1 표기 getRange는 지원하지 않습니다: ' + row);
    return new Range(this, row, col, nr, nc);
  }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  appendRow(arr) { this.getRange(this.getLastRow() + 1, 1, 1, arr.length).setValues([arr]); return this; }
  deleteRow(r) {
    this._grid().splice(r - 1, 1);
    this.rows--;
    this.ss._op({ req: { deleteDimension: { range: { sheetId: this.id, dimension: 'ROWS', startIndex: r - 1, endIndex: r } } } });
    return this;
  }
  clearContents() {
    this.grid = [];
    this.ss._op({ req: { updateCells: { range: { sheetId: this.id }, fields: 'userEnteredValue' } } });
    return this;
  }
  setFrozenRows(n) {
    this.ss._op({ req: { updateSheetProperties: { properties: { sheetId: this.id, gridProperties: { frozenRowCount: n } }, fields: 'gridProperties.frozenRowCount' } } });
    return this;
  }
  setColumnWidths(start, num, px) {
    this.ss._op({ req: { updateDimensionProperties: { range: { sheetId: this.id, dimension: 'COLUMNS', startIndex: start - 1, endIndex: start - 1 + num }, properties: { pixelSize: px }, fields: 'pixelSize' } } });
    return this;
  }
  activate() { return this; }
}

class Spreadsheet {
  constructor(id, token, DateCtor) {
    this.id = id; this.token = token; this.ops = []; this.Date = DateCtor || Date;
    let c = cacheGet(id);
    if (!c) {
      c = { at: Date.now(), meta: google(token, 'GET', SHEETS_API + id + '?fields=' + encodeURIComponent('sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))')) };
      sheetCache.set(id, c);
    }
    this.sheets = (c.meta.sheets || []).map(s => new Sheet(this, JSON.parse(JSON.stringify(s.properties))));
  }
  _op(op) { this.ops.push(op); }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(name) { return this._insertSheet(name, false); }
  _insertSheet(name, hidden) {
    if (this.getSheetByName(name)) throw new Error('이미 같은 이름의 시트가 있습니다: ' + name);
    const props = { sheetId: Math.floor(Math.random() * 2e9), title: name, gridProperties: { rowCount: 1000, columnCount: 26 }, hidden: !!hidden };
    this._op({ req: { addSheet: { properties: props } } });
    const s = new Sheet(this, props, []);
    this.sheets.push(s);
    return s;
  }
  setActiveSheet(s) { return s; }
  // 작업 이력 한 줄 추가 — values:append라 기존 행을 읽지 않는다(이력이 길어져도 읽기 한도와 무관).
  appendAudit(row) {
    if (!this.getSheetByName(AUDIT_SHEET)) {
      google(this.token, 'POST', SHEETS_API + this.id + ':batchUpdate', { requests: [{ addSheet: { properties: { title: AUDIT_SHEET, gridProperties: { frozenRowCount: 1 } } } }] });
      cacheDrop(this.id);
      row = [['일시', '사용자', '기능', '내용'], row];
    } else row = [row];
    google(this.token, 'POST', SHEETS_API + this.id + '/values/' + encodeURIComponent("'" + AUDIT_SHEET + "'!A:D") + ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS', { values: row });
  }
  // GAS에는 없는 확장: 여러 시트를 values:batchGet 한 쌍(2회)으로 미리 읽는다(시트마다 2회씩 읽으면 분당 한도에 금방 닿음).
  // .gs 쪽에서는 if (ss._prefetch) ss._prefetch([...]) 로만 부른다(GAS에서는 그냥 건너뜀).
  _prefetch(names) {
    const todo = names.map(n => this.getSheetByName(n)).filter(sh => sh && !sh.grid && !(isCacheable(sh.name) && cacheGet(this.id + ' ' + sh.name)));
    if (todo.length < 2) return;
    const base = SHEETS_API + this.id + '/values:batchGet?' + todo.map(sh => 'ranges=' + encodeURIComponent("'" + sh.name.replace(/'/g, "''") + "'")).join('&')
      + '&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=';
    const [serial, text] = httpRetry(['SERIAL_NUMBER', 'FORMATTED_STRING'].map(o => ({ method: 'GET', url: base + o, headers: { Authorization: 'Bearer ' + this.token } })))
      .map(r => {
        if (r.code !== 200) throw new Error('시트 일괄 읽기 실패 → ' + r.code + ' ' + r.body.toString('utf8').slice(0, 300));
        return JSON.parse(r.body.toString('utf8')).valueRanges || [];
      });
    todo.forEach((sh, i) => {
      const c = { at: Date.now(), serial: (serial[i] || {}).values || [], text: (text[i] || {}).values || [] };
      if (isCacheable(sh.name)) sheetCache.set(this.id + ' ' + sh.name, c);
      sh.grid = c.serial.map((row, ri) => row.map((v, j) => cellValue(v, (c.text[ri] || [])[j], this.Date)));
    });
  }
  // 같은 종류(값 쓰기 / 구조 변경)끼리 연속된 작업을 한 번의 API 호출로 묶되, 전체 순서는 지킨다.
  flush() {
    const ops = this.ops; this.ops = [];
    for (let i = 0; i < ops.length;) {
      const kind = ops[i].values ? 'values' : 'req';
      const group = [];
      while (i < ops.length && (ops[i].values ? 'values' : 'req') === kind) group.push(ops[i++][kind]);
      if (kind === 'values') {
        group.forEach(d => cacheDrop(this.id, d.range.slice(1, d.range.lastIndexOf("'!")).replace(/''/g, "'")));
        google(this.token, 'POST', SHEETS_API + this.id + '/values:batchUpdate', { valueInputOption: 'USER_ENTERED', data: group });
      } else {
        cacheDrop(this.id); // 행 삭제·시트 추가 등 구조 변경은 범위를 따지지 않고 이 스프레드시트 캐시를 통째로 버린다
        google(this.token, 'POST', SHEETS_API + this.id + ':batchUpdate', { requests: group });
      }
    }
  }
}

// ---------- DriveApp ----------
const FILE_FIELDS = 'id,name,mimeType,webViewLink,size';
const driveQ = s => "'" + String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };

class DriveFile {
  constructor(rt, meta) { this.rt = rt; this.meta = meta; }
  _m() { if (!this.meta.mimeType) this.meta = this.rt.drive('GET', DRIVE_API + '/' + this.meta.id + '?supportsAllDrives=true&fields=' + FILE_FIELDS); return this.meta; }
  getId() { return this.meta.id; }
  getName() { return this._m().name; }
  getMimeType() { return this._m().mimeType; }
  getSize() { return Number(this._m().size || 0); }
  getUrl() { return this._m().webViewLink; }
  getBlob() {
    const r = httpAll([{ method: 'GET', url: DRIVE_API + '/' + this.meta.id + '?alt=media&supportsAllDrives=true', headers: { Authorization: 'Bearer ' + this.rt.token } }])[0];
    if (r.code !== 200) throw new Error('Drive 파일을 읽을 수 없습니다(' + r.code + '): ' + this.meta.id);
    return new Blob(r.body, this.meta.mimeType || (r.type || '').split(';')[0], this.meta.name);
  }
  setSharing() {
    this.rt.drive('POST', DRIVE_API + '/' + this.meta.id + '/permissions?supportsAllDrives=true', { role: 'reader', type: 'anyone' });
    return this;
  }
}

class DriveFolder extends DriveFile {
  _children(q) {
    const out = [];
    let pageToken = '';
    do {
      const res = this.rt.drive('GET', DRIVE_API + '?pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=' + encodeURIComponent('nextPageToken,files(' + FILE_FIELDS + ')')
        + '&q=' + encodeURIComponent(driveQ(this.meta.id) + ' in parents and trashed=false and ' + q) + (pageToken ? '&pageToken=' + pageToken : ''));
      out.push(...res.files);
      pageToken = res.nextPageToken;
    } while (pageToken);
    return out;
  }
  getFiles() { return iter(this._children("mimeType!='application/vnd.google-apps.folder'").map(m => new DriveFile(this.rt, m))); }
  getFoldersByName(name) { return iter(this._children("mimeType='application/vnd.google-apps.folder' and name=" + driveQ(name)).map(m => new DriveFolder(this.rt, m))); }
  createFolder(name) { return new DriveFolder(this.rt, this.rt.drive('POST', DRIVE_API + '?supportsAllDrives=true&fields=' + FILE_FIELDS, { name, mimeType: 'application/vnd.google-apps.folder', parents: [this.meta.id] })); }
  createFile(blob) {
    const boundary = 'gasshim' + Date.now();
    const head = Buffer.from('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'
      + JSON.stringify({ name: blob.getName() || 'untitled', parents: [this.meta.id] })
      + '\r\n--' + boundary + '\r\nContent-Type: ' + blob.getContentType() + '\r\n\r\n');
    const body = Buffer.concat([head, blob.getBytes(), Buffer.from('\r\n--' + boundary + '--')]);
    const r = httpAll([{ method: 'POST', url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=' + FILE_FIELDS,
      headers: { Authorization: 'Bearer ' + this.rt.token, 'Content-Type': 'multipart/related; boundary=' + boundary }, body64: body.toString('base64') }])[0];
    if (r.code !== 200) throw new Error('Drive 업로드 실패(' + r.code + '): ' + r.body.toString('utf8').slice(0, 300));
    return new DriveFile(this.rt, JSON.parse(r.body.toString('utf8')));
  }
}

// ---------- 요청 단위 런타임 ----------
// 요청마다 새 vm 컨텍스트를 만든다 — 스프레드시트 캐시/쓰기 버퍼가 요청 사이에 섞이지 않게.
function createRuntime({ token, appUrl, user, audit }) {
  const opened = {};
  const triggers = [];
  const rt = { token, drive: (m, url, json) => google(token, m, url, json) };

  const SpreadsheetApp = {
    openById: id => opened[id] || (opened[id] = new Spreadsheet(id, token, vm.runInContext('Date', ctx))),
    getUi: () => { throw new Error('이 기능은 구글시트 상단 메뉴에서 실행하세요.'); },
    flush: () => {}
  };

  // Script Properties는 숨김 시트(_properties)에 key/value로 저장한다. 없는 키는 같은 이름의 환경변수로 대체.
  const propSheet = () => {
    const ss = SpreadsheetApp.openById(ctx.SHEET_ID);
    return ss.getSheetByName(PROPS_SHEET) || ss._insertSheet(PROPS_SHEET, true);
  };
  const propRows = () => { const s = propSheet(); return s.getLastRow() ? s.getRange(1, 1, s.getLastRow(), 2).getValues() : []; };
  const props = {
    getProperty: k => {
      const r = propRows().find(x => x[0] === k);
      return r && !isEmpty(r[1]) ? String(r[1]) : (process.env[k] || null);
    },
    setProperty: (k, v) => {
      const i = propRows().findIndex(x => x[0] === k);
      const s = propSheet();
      if (i >= 0) s.getRange(i + 1, 2).setValue(String(v)); else s.appendRow([k, String(v)]);
      return props;
    },
    deleteProperty: k => { const i = propRows().findIndex(x => x[0] === k); if (i >= 0) propSheet().deleteRow(i + 1); return props; },
    getProperties: () => Object.fromEntries(propRows().filter(r => r[0]).map(r => [r[0], String(r[1])]))
  };

  // ponytail: 잠금은 no-op. 검수 배치는 하나의 체인으로만 이어지므로 실사용상 충돌이 드묾. 중복 처리가 보이면 _properties 기반 잠금 추가.
  const lock = { waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true };

  const ctx = vm.createContext({
    console,
    Logger: { log: (...a) => console.log(...a) },
    Utilities,
    UrlFetchApp,
    SpreadsheetApp,
    PropertiesService: { getScriptProperties: () => props, getDocumentProperties: () => props },
    LockService: { getScriptLock: () => lock, getDocumentLock: () => lock },
    Session: { getScriptTimeZone: () => TZ, getEffectiveUser: () => ({ getEmail: () => user }), getActiveUser: () => ({ getEmail: () => user }) },
    ScriptApp: {
      getOAuthToken: () => token,
      getService: () => ({ getUrl: () => appUrl }),
      getProjectTriggers: () => [],
      deleteTrigger() {},
      newTrigger: fn => ({ timeBased: () => ({ after: ms => ({ create: () => { triggers.push({ fn, ms }); } }) }) })
    },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
      Permission: { VIEW: 'VIEW' },
      getFileById: id => new DriveFile(rt, { id }),
      getFolderById: id => { const f = new DriveFolder(rt, { id }); f._m(); return f; },
      getRootFolder: () => new DriveFolder(rt, { id: 'root', mimeType: 'application/vnd.google-apps.folder' }),
      getFoldersByName: name => DriveApp().getRootFolder().getFoldersByName(name),
      createFolder: name => DriveApp().getRootFolder().createFolder(name)
    }
  });
  const DriveApp = () => ctx.DriveApp;
  scripts.forEach(s => s.runInContext(ctx));
  ctx.SHEET_ID = vm.runInContext('SHEET_ID', ctx);

  return {
    ctx,
    triggers,
    // GAS처럼 함수가 중간에 예외로 끝나도 그 전까지의 시트 쓰기는 반영한다.
    // audit이면 실제로 시트에 뭔가 쓴 호출만 작업이력 시트에 남긴다(누가·언제·무엇을).
    call(fn, args) {
      try {
        return ctx[fn].apply(null, args || []);
      } finally {
        const wrote = Object.values(opened).some(ss => ss.ops.length);
        Object.values(opened).forEach(ss => ss.flush());
        if (audit && wrote) {
          try {
            SpreadsheetApp.openById(ctx.SHEET_ID).appendAudit([formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), user, fn, JSON.stringify(args || []).slice(0, 500)]);
          } catch (e) { console.error('작업이력 기록 실패', e.message); }
        }
      }
    }
  };
}

module.exports = { createRuntime, formatDate, Spreadsheet, Sheet, httpStats };
