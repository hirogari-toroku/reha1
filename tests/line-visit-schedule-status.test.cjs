const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const c = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.getScheduleWindowReadStartRow_ = () => 2;
  return c;
}

function makeSheet(rows) {
  const writes = [];
  return {
    rows,
    writes,
    getLastRow: () => rows.length,
    getLastColumn: () => 11,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = (rows[row - 1 + i] || [])[col - 1 + j];
          return v === undefined ? '' : v;
        })),
      setValues: values => {
        writes.push({ row, col });
        values.forEach((line, i) => line.forEach((v, j) => { rows[row - 1 + i][col - 1 + j] = v; }));
      }
    })
  };
}

const header = ['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', '', '状態', '更新日時', '更新内容'];
const received = new Date(2026, 8, 22, 10, 2);
const schedule = (staff, user, date, status) => [new Date(2026, 8, 1), staff, user, date, '', '', '', '', status, '', ''];

test('LINE 開始 moves the single matching schedule to 訪問中 and 終了 then completes it', () => {
  const c = context();
  const sheet = makeSheet([header, schedule('佐藤', '山田太郎', '2026-09-21', '予定'), schedule('佐藤', '山田太郎', '2026-09-22', '予定')]);
  assert.equal(c.updateScheduleStatusFromLineVisit_(sheet, '佐藤', '山田太郎', '開始', '9/22', '10:02', received), 3);
  assert.equal(sheet.rows[2][8], '訪問中');
  assert.equal(sheet.rows[1][8], '予定', 'other days are untouched');
  assert.equal(c.updateScheduleStatusFromLineVisit_(sheet, '佐藤', '山田太郎', '終了', '9/22', '11:00', received), 3);
  assert.equal(sheet.rows[2][8], '完了');
});

test('ambiguous, cancelled, completed, or other-staff schedules are left alone', () => {
  const c = context();
  const twoPlanned = makeSheet([header, schedule('佐藤', '山田太郎', '2026-09-22', '予定'), schedule('佐藤', '山田太郎', '2026-09-22', '')]);
  assert.equal(c.updateScheduleStatusFromLineVisit_(twoPlanned, '佐藤', '山田太郎', '開始', '9/22', '10:02', received), null);
  assert.equal(twoPlanned.writes.length, 0);

  const others = makeSheet([
    header,
    schedule('佐藤', '山田太郎', '2026-09-22', 'キャンセル'),
    schedule('佐藤', '山田太郎', '2026-09-22', '完了'),
    schedule('田中', '山田太郎', '2026-09-22', '予定')
  ]);
  assert.equal(c.updateScheduleStatusFromLineVisit_(others, '佐藤', '山田太郎', '開始', '9/22', '10:02', received), null);
  assert.equal(others.writes.length, 0);
});

test('a sheet error never throws into the LINE visit flow', () => {
  const c = context();
  const broken = { getLastRow: () => { throw new Error('boom'); } };
  c.console = { error() {} };
  assert.equal(c.updateScheduleStatusFromLineVisit_(broken, '佐藤', '山田太郎', '開始', '9/22', '10:02', received), null);
});

test('LIFF 終了 is allowed when the start was logged through LINE even if the row is still 予定', () => {
  const src = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  assert.match(src, /type === "終了" && scheduleStatus !== "訪問中" && !hasLoggedVisitStart_\(/);
  const c = context();
  c.buildVisitStatusIndex_ = () => ({});
  c.getVisitStatusForSchedule_ = () => ({ hasStart: true, hasEnd: false });
  assert.equal(c.hasLoggedVisitStart_({}, '佐藤', '山田太郎', '2026-09-22', received), true);
  c.getVisitStatusForSchedule_ = () => null;
  assert.equal(c.hasLoggedVisitStart_({}, '佐藤', '山田太郎', '2026-09-22', received), false);
});
