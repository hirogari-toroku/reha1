const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.Utilities = {
    formatDate: (date, _tz, format) => {
      const pad = n => String(n).padStart(2, '0');
      if (format === 'HH:mm') return pad(date.getHours()) + ':' + pad(date.getMinutes());
      return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate());
    }
  };
  return c;
}

function makeSheet(rows) {
  return {
    getLastRow: () => rows.length,
    getLastColumn: () => 8,
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = (rows[row - 1 + i] || [])[col - 1 + j];
          return v === undefined ? '' : v;
        }))
    })
  };
}

const header = ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻', '', ''];
const recordedHoursAgo = h => new Date(Date.now() - h * 3600 * 1000);

test('an end already recorded hours ago through the other channel at a nearby time is a duplicate', () => {
  const c = context();
  const sheet = makeSheet([header, [recordedHoursAgo(3), '終了', '佐藤', '山田太郎', '2026-09-22', '15:00']]);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '終了', '2026-09-22', '15:20', new Date()), true);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '終了', '2026/9/22', '14:05', new Date()), true);
});

test('a second visit later the same day, another type, user, staff, or date is not a duplicate', () => {
  const c = context();
  const sheet = makeSheet([header, [recordedHoursAgo(3), '終了', '佐藤', '山田太郎', '2026-09-22', '10:00']]);
  const now = new Date();
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '終了', '2026-09-22', '15:00', now), false);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '開始', '2026-09-22', '10:00', now), false);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '鈴木花子', '終了', '2026-09-22', '10:00', now), false);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '田中', '山田太郎', '終了', '2026-09-22', '10:00', now), false);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '終了', '2026-09-23', '10:00', now), false);
});

test('records without a usable time fall back to the original exact-match rule only', () => {
  const c = context();
  const sheet = makeSheet([header, [recordedHoursAgo(3), '終了', '佐藤', '山田太郎', '2026-09-22', '']]);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '終了', '2026-09-22', '15:00', new Date()), false);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '終了', '2026-09-22', '', new Date()), false);
});

test('the original same-minute resend check still applies', () => {
  const c = context();
  const now = new Date();
  const sheet = makeSheet([header, [recordedHoursAgo(0.05), '開始', '佐藤', '山田太郎', '2026-09-22', '09:00']]);
  assert.equal(c.isRecentDuplicateVisit_(sheet, '佐藤', '山田太郎', '開始', '2026-09-22', '09:00', now), true);
});
