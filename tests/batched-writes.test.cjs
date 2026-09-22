const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  return c;
}

// Grid-backed fake sheet that records every Range write so tests can count calls.
function makeSheet(initialRows) {
  const rows = initialRows.map(r => r.slice());
  const writes = [];
  const cell = (r, c) => (rows[r - 1] || [])[c - 1];
  const put = (r, c, v) => {
    while (rows.length < r) rows.push([]);
    rows[r - 1][c - 1] = v;
  };
  return {
    rows,
    writes,
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((max, r) => Math.max(max, r.length), 0),
    getDataRange() {
      return this.getRange(1, 1, rows.length, this.getLastColumn());
    },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValue: () => cell(row, col),
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = cell(row + i, col + j);
          return v === undefined ? '' : v;
        })),
      setValue: v => { writes.push({ row, col, numRows: 1, numCols: 1 }); put(row, col, v); },
      setValues: values => {
        assert.equal(values.length, numRows);
        values.forEach(line => assert.equal(line.length, numCols));
        writes.push({ row, col, numRows, numCols });
        values.forEach((line, i) => line.forEach((v, j) => put(row + i, col + j, v)));
      }
    })
  };
}

test('applyCellUpdates_ merges neighbouring cells into blocks and writes only listed cells', () => {
  const c = context();
  const sheet = makeSheet([['h1', 'h2', 'h3', 'h4'], ['a', '=SUM(1)', 'x', ''], ['b', 'keep', 'y', ''], ['c', '', 'z', '']]);
  const blocks = c.applyCellUpdates_(sheet, [
    { rowNumber: 2, column: 3, value: 'X' },
    { rowNumber: 3, column: 3, value: 'Y' },
    { rowNumber: 4, column: 3, value: 'Z' },
    { rowNumber: 2, column: 4, value: 'D2' },
    { rowNumber: 3, column: 4, value: 'D3' },
    { rowNumber: 4, column: 4, value: 'D4' }
  ]);
  assert.equal(blocks, 1);
  assert.deepEqual(sheet.writes, [{ row: 2, col: 3, numRows: 3, numCols: 2 }]);
  assert.equal(sheet.rows[1][1], '=SUM(1)');
  assert.deepEqual(sheet.rows.map(r => r[2]), ['h3', 'X', 'Y', 'Z']);
});

test('applyCellUpdates_ keeps gaps unwritten and lets a later update to the same cell win', () => {
  const c = context();
  const sheet = makeSheet([['h'], ['1', '', ''], ['2', 'keep', ''], ['3', '', '']]);
  c.applyCellUpdates_(sheet, [
    { rowNumber: 2, column: 2, value: 'first' },
    { rowNumber: 2, column: 2, value: 'second' },
    { rowNumber: 4, column: 2, value: 'four' }
  ]);
  assert.equal(sheet.writes.length, 2);
  assert.equal(sheet.rows[1][1], 'second');
  assert.equal(sheet.rows[2][1], 'keep');
  assert.equal(sheet.rows[3][1], 'four');
  assert.equal(c.applyCellUpdates_(sheet, []), 0);
});

test('payment reconciliation writes the changed G:H span in one call and skips unchanged runs', () => {
  const c = context();
  const header = ['日付', '摘要', '入金', '出金', '残高', '備考', '照合利用者', '照合状態'];
  const payments = makeSheet([
    header,
    ['2026/09/01', '振込 ヤマダ タロウ', 30000, '', '', '', '', ''],
    ['2026/09/02', '利息', 700, '', '', '', 'manual', 'memo'],
    ['2026/09/03', '振込 スズキ ハナコ', 60000, '', '', '', '', ''],
    ['2026/09/04', 'デビット', '', 500, '', '', '', '']
  ]);
  const users = makeSheet([['利用者名', 'カナ1', 'カナ2', 'カナ3'], ['山田太郎', 'ヤマダタロウ', '', ''], ['鈴木花子', 'スズキハナコ', '', '']]);
  const ss = { getSheetByName: name => (name === '入出金明細' ? payments : name === '利用者マスタ' ? users : null) };

  assert.equal(c.reconcileNyushukkinUsersCore_(ss), 2);
  assert.deepEqual(payments.writes, [{ row: 2, col: 7, numRows: 3, numCols: 2 }]);
  assert.deepEqual(payments.rows[1].slice(6), ['山田太郎', '照合済']);
  assert.deepEqual(payments.rows[2].slice(6), ['manual', 'memo'], 'unmatched row inside the span keeps its values');
  assert.deepEqual(payments.rows[3].slice(6), ['鈴木花子', '照合済']);

  payments.writes.length = 0;
  assert.equal(c.reconcileNyushukkinUsersCore_(ss), 2, 'count still reports matched rows');
  assert.equal(payments.writes.length, 0, 'already reconciled rows are not rewritten');
});

test('schedule status headers are read once and written only when missing', () => {
  const c = context();
  const complete = makeSheet([['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', '状態', '更新日時', '更新内容']]);
  c.ensureScheduleStatusColumns_(complete);
  assert.equal(complete.writes.length, 0);

  const partial = makeSheet([['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', '状態', '', '']]);
  c.ensureScheduleStatusColumns_(partial);
  assert.deepEqual(partial.writes, [{ row: 1, col: 9, numRows: 1, numCols: 3 }]);
  assert.deepEqual(partial.rows[0].slice(8), ['状態', '更新日時', '更新内容']);
});

test('fillMissingHeaders_ only fills blank header cells', () => {
  const c = context();
  const sheet = makeSheet([['A', '', 'custom']]);
  c.fillMissingHeaders_(sheet, ['A', 'B', 'C', 'D']);
  assert.deepEqual(sheet.rows[0], ['A', 'B', 'custom', 'D']);
  assert.equal(sheet.writes.length, 2);
});
