const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context(store) {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: key => (key in store ? store[key] : null),
      setProperty: (key, value) => { store[key] = value; },
      deleteProperty: key => { delete store[key]; }
    })
  };
  c.Utilities = {
    formatDate: (date, _tz, format) => {
      const pad = n => String(n).padStart(2, '0');
      const ymd = date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate());
      if (format === 'yyyy-MM-dd') return ymd;
      if (format === 'yyyy/MM/dd') return ymd.replace(/-/g, '/');
      return ymd + ' ' + pad(date.getHours()) + ':' + pad(date.getMinutes());
    }
  };
  return c;
}

function day(offset) {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + offset);
  return d;
}

function makeSheet(rows) {
  const reads = [];
  return {
    reads,
    getLastRow: () => rows.length,
    getLastColumn: () => 11,
    getRange: (row, col, numRows = 1, numCols = 1) => {
      reads.push({ row, numRows });
      return {
        getValues: () => Array.from({ length: numRows }, (_, i) =>
          Array.from({ length: numCols }, (_, j) => {
            const v = (rows[row - 1 + i] || [])[col - 1 + j];
            return v === undefined ? '' : v;
          }))
      };
    }
  };
}

// Header + 200 old rows (6 months ago) + a far-future booking made long ago in the middle
// + recent rows. The far-future row sits among old rows, which is why the hint must be
// the first in-range row rather than a backward-scan cutoff.
function scheduleRows() {
  const rows = [['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', '', '状態', '更新日時', '更新内容']];
  for (let i = 0; i < 200; i++) rows.push([day(-190), 'S', 'U', day(-180), '', '', '', '', '完了', '', '']);
  rows.push([day(-190), 'S', 'U', day(20), '', '', '', '', '予定', '', '']); // row 202
  for (let i = 0; i < 50; i++) rows.push([day(-190), 'S', 'U', day(-170), '', '', '', '', '完了', '', '']);
  for (let i = 0; i < 5; i++) rows.push([day(-3 + i), 'S', 'U', day(i), '', '', '', '', '予定', '', '']);
  return rows;
}

test('hint is the first row inside the window even when later rows are older', () => {
  const store = {};
  const c = context(store);
  const sheet = makeSheet(scheduleRows());
  assert.equal(c.getScheduleWindowReadStartRow_(sheet), 202);
  const saved = JSON.parse(store.SCHEDULE_READ_START_ROW_HINT);
  assert.equal(saved.row, 202);
  assert.equal(saved.lastRow, 257);
});

test('same-day reads reuse the hint without rescanning; a new day or shrunk sheet rescans', () => {
  const store = {};
  const c = context(store);
  const sheet = makeSheet(scheduleRows());
  c.getScheduleWindowReadStartRow_(sheet);
  sheet.reads.length = 0;
  assert.equal(c.getScheduleWindowReadStartRow_(sheet), 202);
  assert.equal(sheet.reads.length, 0);

  store.SCHEDULE_READ_START_ROW_HINT = JSON.stringify({ day: '2000-1-1', row: 250, lastRow: 257 });
  assert.equal(c.getScheduleWindowReadStartRow_(sheet), 202);

  const today = c.formatScheduleHintDayKey_(new Date());
  store.SCHEDULE_READ_START_ROW_HINT = JSON.stringify({ day: today, row: 250, lastRow: 999 });
  assert.equal(c.getScheduleWindowReadStartRow_(sheet), 202, 'rows were deleted since the hint, so rescan');
});

test('staff and user schedule lists are identical with and without the hint', () => {
  const rows = scheduleRows();
  const withHint = context({});
  const full = context({});
  full.getScheduleWindowReadStartRow_ = () => 2;

  const hinted = withHint.collectActiveSchedulesForStaff_(makeSheet(rows), 'S', {});
  const complete = full.collectActiveSchedulesForStaff_(makeSheet(rows), 'S', {});
  assert.deepEqual(JSON.parse(JSON.stringify(hinted)), JSON.parse(JSON.stringify(complete)));
  assert.ok(hinted.some(s => s.id === '202'), 'row ids stay the real sheet row numbers');
  assert.equal(hinted.length, 6);

  const userHinted = withHint.collectActiveSchedulesForUser_(makeSheet(rows), 'U', {});
  const userComplete = full.collectActiveSchedulesForUser_(makeSheet(rows), 'U', {});
  assert.deepEqual(JSON.parse(JSON.stringify(userHinted)), JSON.parse(JSON.stringify(userComplete)));
});

test('moving an old row into the window lowers the saved hint', () => {
  const today = new Date();
  const store = { SCHEDULE_READ_START_ROW_HINT: JSON.stringify({ day: 'x', row: 202, lastRow: 257 }) };
  const c = context(store);
  c.lowerScheduleReadStartRowHint_(10);
  assert.equal(JSON.parse(store.SCHEDULE_READ_START_ROW_HINT).row, 10);
  c.lowerScheduleReadStartRowHint_(300);
  assert.equal(JSON.parse(store.SCHEDULE_READ_START_ROW_HINT).row, 10);
  assert.ok(today);
});

test('an unparseable visit date keeps the row in range like the old full read', () => {
  const c = context({});
  const rows = [['h'], [day(-190), 'S', 'U', day(-180)], [day(-190), 'S', 'U', 'そのうち'], [day(-1), 'S', 'U', day(1)]];
  assert.equal(c.getScheduleWindowReadStartRow_(makeSheet(rows)), 3);
});
