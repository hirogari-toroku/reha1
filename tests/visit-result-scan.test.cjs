const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function makeSheet(initialRows) {
  const rows = initialRows.map(r => r.slice());
  return {
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((max, r) => Math.max(max, r.length), 0),
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => {
        const out = [];
        for (let i = 0; i < numRows; i++) {
          const source = rows[row - 1 + i] || [];
          const line = [];
          for (let j = 0; j < numCols; j++) line.push(source[col - 1 + j]);
          out.push(line);
        }
        return out;
      }
    }),
    rows
  };
}

function context() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.Utilities = {
    formatDate: (date, _tz, format) => {
      const pad = n => String(n).padStart(2, '0');
      if (format === 'yyyy-MM-dd') return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate());
      return String(date);
    }
  };
  return c;
}

function day(offsetFromToday) {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + offsetFromToday);
  return d;
}

test('without earliestDate the full sheet is read from row 2 (unchanged legacy behavior)', () => {
  const c = context();
  const header = ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻'];
  const rows = [header];
  for (let i = 0; i < 10; i++) rows.push([day(-500 + i), '開始', 'S', 'A', '', '']);
  const sheet = makeSheet(rows);
  assert.equal(c.findVisitResultReadStartRow_(sheet, undefined), 2);
});

test('scan stops once rows are older than earliestDate minus the safety margin, across multiple chunks', () => {
  const c = context();
  c.VISIT_RESULT_SCAN_CHUNK_ROWS = 5; // force multiple chunks over a small dataset
  const header = ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻'];
  const rows = [header];
  // rows 2..13: old visits far outside the window (should be skipped)
  for (let i = 0; i < 12; i++) rows.push([day(-400 + i), '開始', 'S', 'A', '', '']);
  // rows 14..18: recent visits inside the window (must be read)
  for (let i = 0; i < 5; i++) rows.push([day(-10 + i), '開始', 'S', 'A', '', '']);
  const sheet = makeSheet(rows);

  const earliestDate = day(-14); // matches the display window's start
  const startRow = c.findVisitResultReadStartRow_(sheet, earliestDate);

  // Every row from startRow onward must be within (or safely before) the window.
  for (let row = startRow; row <= sheet.getLastRow(); row++) {
    const cellDate = sheet.rows[row - 1][0];
    assert.ok(cellDate.getTime() >= earliestDate.getTime() - 31 * 24 * 60 * 60 * 1000,
      'row ' + row + ' should not have been skipped');
  }
  // The recent, in-window rows must all be included.
  assert.ok(startRow <= 14, '直近の予定が読み取り範囲から漏れてはいけない');
});

test('a safety-margin row just before earliestDate is still included, not treated as old', () => {
  const c = context();
  const header = ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻'];
  const earliestDate = day(-60);
  const withinSafetyMargin = new Date(earliestDate.getTime() - 10 * 24 * 60 * 60 * 1000); // 10 days before window start
  const trulyOld = new Date(earliestDate.getTime() - 40 * 24 * 60 * 60 * 1000); // 40 days before window start
  const rows = [header,
    [trulyOld, '開始', 'S', 'A', '', ''],
    [withinSafetyMargin, '開始', 'S', 'A', '', ''],
    [day(0), '開始', 'S', 'A', '', '']
  ];
  const sheet = makeSheet(rows);
  const startRow = c.findVisitResultReadStartRow_(sheet, earliestDate);
  assert.ok(startRow <= 3, '安全余裕(30日)以内の行は読み取り対象から漏れてはいけない');
});

test('blank or malformed date cells do not break the backward scan', () => {
  const c = context();
  const header = ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻'];
  const rows = [header,
    ['', '開始', 'S', 'A', '', ''],
    [day(-1), '開始', 'S', 'A', '', '']
  ];
  const sheet = makeSheet(rows);
  assert.doesNotThrow(() => c.findVisitResultReadStartRow_(sheet, day(0)));
});

test('buildVisitStatusIndex_ still finds in-window entries after bounding the read', () => {
  const c = context();
  const header = ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻', '本文', 'LINEユーザーID'];
  const rows = [header];
  for (let i = 0; i < 5; i++) rows.push([day(-300 + i), '開始', 'S', 'OLD', '1/1', '10:00', '', '']);
  const recentDate = day(-5);
  const dateText = (recentDate.getMonth() + 1) + '/' + recentDate.getDate();
  rows.push([recentDate, '開始', 'S', 'A', dateText, '10:00', '', '']);
  const sheet = makeSheet(rows);

  const index = c.buildVisitStatusIndex_(sheet, 'S', 'A', day(-60));
  const key = Object.keys(index).find(k => k.endsWith('|s|a') || index[k].userName === 'A');
  assert.ok(key, '直近の実績が結果に含まれるべき');
  assert.equal(index[key].hasStart, true);
});
