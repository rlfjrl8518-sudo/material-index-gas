// node lib/gas.test.js — HTTP 없이 시트 grid/쓰기 버퍼 동작만 확인한다.
const assert = require('assert');
const { Spreadsheet, Sheet } = require('./gas');

const ss = Object.assign(Object.create(Spreadsheet.prototype), { id: 'S', token: 't', ops: [] });
const sh = new Sheet(ss, { sheetId: 7, title: "시트'1", gridProperties: { rowCount: 3, columnCount: 2 } }, [['h1', 'h2'], ['a', 1]]);
ss.sheets = [sh];

assert.equal(sh.getLastRow(), 2);
sh.appendRow(['b', new Date('2026-10-27T15:00:00Z'), 'x']);
assert.equal(sh.getLastRow(), 3);
assert.equal(sh.getLastColumn(), 3);
assert.deepEqual(sh.getRange(2, 1, 2, 1).getValues(), [['a'], ['b']]);
assert.deepEqual(sh.getRange(5, 1).getValues(), [['']]);
sh.deleteRow(2);
assert.deepEqual(sh.getDataRange().getValues()[1].slice(0, 1), ['b']);
assert.throws(() => sh.getRange(1, 1, 2, 2).setValues([[1, 2]]));

// appendRow가 열 개수(2)를 넘겨서 열 확장 요청이 값 쓰기보다 먼저 들어가야 한다
assert.deepEqual(ss.ops.map(o => Object.keys(o.req || o)[0]), ['appendDimension', 'values', 'deleteDimension']);
const w = ss.ops[1].values;
assert.equal(w.range, "'시트''1'!A3:C3");
assert.deepEqual(w.values, [['b', '2026-10-28 00:00:00', 'x']]);
// 다른 realm(vm)의 Date도 날짜로 써져야 한다
const vmDate = require('vm').runInNewContext('new Date("2026-10-27T15:00:00Z")');
sh.getRange(4, 1).setValue(vmDate);
assert.deepEqual(ss.ops[ss.ops.length - 1].values.values, [['2026-10-28 00:00:00']]);
console.log('gas.test ok');
