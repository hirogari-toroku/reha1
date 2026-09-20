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
      getValue: () => (rows[row - 1] || [])[col - 1],
      getValues: () => {
        const out = [];
        for (let i = 0; i < numRows; i++) {
          const source = rows[row - 1 + i] || [];
          const line = [];
          for (let j = 0; j < numCols; j++) line.push(source[col - 1 + j]);
          out.push(line);
        }
        return out;
      },
      setValue: value => { rows[row - 1] = rows[row - 1] || []; rows[row - 1][col - 1] = value; },
      setValues: values => {
        values.forEach((line, i) => {
          rows[row - 1 + i] = rows[row - 1 + i] || [];
          line.forEach((value, j) => { rows[row - 1 + i][col - 1 + j] = value; });
        });
      }
    }),
    appendRow: row => rows.push(row.slice()),
    getDataRange: () => ({ getValues: () => rows }),
    rows
  };
}

function context() {
  const c = vm.createContext({ console: { error() {} } });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.Utilities = { formatDate: (date, _tz, format) => {
    const pad = n => String(n).padStart(2, '0');
    const map = { M: String(date.getMonth() + 1), d: String(date.getDate()), H: String(date.getHours()), mm: pad(date.getMinutes()) };
    return format.replace(/M|d|H|mm/g, t => map[t]);
  } };
  c.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  c.getStaffName_ = () => 'S';
  c.resolveUserName_ = () => 'A';
  c.isAmbiguousUserName_ = () => false;
  c.isRegisteredStaffUser_ = () => true;
  c.isDuplicateSchedule_ = () => false;
  c.saveUnknownUser_ = () => {};
  c.saveLiffOperationLog_ = () => {};
  c.sendScheduleConfirmationPushFromLiff_ = () => {};
  // CalendarApp is intentionally left undefined: any remaining calendar-sync
  // code path would throw ReferenceError as soon as it is touched.
  return c;
}

test('recordScheduleFromLiff_ registers a schedule without touching CalendarApp', () => {
  const c = context();
  const scheduleSheet = makeSheet([['登録日時', 'スタッフ名', '利用者名', '訪問日']]);
  const ss = { getSheetByName: name => name === '訪問予定' ? scheduleSheet : makeSheet([]) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };

  const result = c.recordScheduleFromLiff_('id', 'A', '9/25');

  assert.equal(result.success, true);
  assert.equal(scheduleSheet.rows.length, 2);
  assert.equal(scheduleSheet.rows[1][2], 'A');
  assert.equal(scheduleSheet.rows[1][3], '9/25');
  assert.equal(scheduleSheet.rows[1][6], '');
  assert.equal(scheduleSheet.rows[1][7], '');
});

test('cancelScheduleFromLiff_ cancels a schedule without touching CalendarApp', () => {
  const c = context();
  const scheduleSheet = makeSheet([
    ['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', 'old-event-id', '予定', '', ''],
    [new Date(), 'S', 'A', new Date(), '', '', '', 'old-event-id', '予定', '', '']
  ]);
  const ss = { getSheetByName: name => name === '訪問予定' ? scheduleSheet : makeSheet([]) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };

  const result = c.cancelScheduleFromLiff_('id', '2');

  assert.equal(result.success, true);
  assert.equal(scheduleSheet.rows[1][8], 'キャンセル');
});

test('updateScheduleFromLiff_ moves a schedule without touching CalendarApp', () => {
  const c = context();
  const scheduleSheet = makeSheet([
    ['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', 'old-event-id', '予定', '', ''],
    [new Date(), 'S', 'A', new Date(), '', '', '', 'old-event-id', '予定', '', '']
  ]);
  const ss = { getSheetByName: name => name === '訪問予定' ? scheduleSheet : makeSheet([]) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };

  const result = c.updateScheduleFromLiff_('id', '2', 'A', '9/26');

  assert.equal(result.success, true);
  assert.equal(scheduleSheet.rows[1][3], '9/26');
  assert.equal(scheduleSheet.rows[1][6], '');
  assert.equal(scheduleSheet.rows[1][7], '');
});

test('LINE text schedule registration (doPost) appends without touching CalendarApp', () => {
  const c = context();
  const scheduleSheet = makeSheet([['登録日時', 'スタッフ名', '利用者名', '訪問日']]);
  const staffUserSheet = makeSheet([['スタッフ名', '利用者名'], ['S', 'A']]);
  const otherSheets = {};
  const ss = {
    getSheetByName: name => {
      if (name === '訪問予定') return scheduleSheet;
      if (name === 'スタッフ利用者マスタ') return staffUserSheet;
      if (!otherSheets[name]) otherSheets[name] = makeSheet([]);
      return otherSheets[name];
    },
    insertSheet: name => { otherSheets[name] = makeSheet([]); return otherSheets[name]; }
  };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };
  c.ContentService = { createTextOutput: () => ({ setMimeType() { return this; } }), MimeType: { TEXT: 'text' } };
  c.sendReplyMessages_ = () => {};
  c.getLineDisplayNameFromEvent_ = () => 'S';
  c.saveLineUserDirectory_ = () => {};
  c.extractDates_ = () => ['9/25'];
  c.parseSchedule_ = () => [[new Date(), 'S', 'A', '9/25', '', 'line-id']];

  c.doPost({ postData: { contents: JSON.stringify({
    events: [{ type: 'message', source: { userId: 'line-id' }, replyToken: 'rt',
      message: { type: 'text', text: '山田太郎\n9/25\nお願いします' } }]
  }) } });

  assert.equal(scheduleSheet.rows.length, 2);
  assert.equal(scheduleSheet.rows[1][2], 'A');
  assert.equal(scheduleSheet.rows[1][6], '');
  assert.equal(scheduleSheet.rows[1][7], '');
});
