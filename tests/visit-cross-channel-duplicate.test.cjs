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
  c.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  c.Utilities = {
    formatDate: (date, _tz, format) => {
      const pad = n => String(n).padStart(2, '0');
      const map = {
        yyyy: String(date.getFullYear()), MM: pad(date.getMonth() + 1), M: String(date.getMonth() + 1),
        dd: pad(date.getDate()), d: String(date.getDate()),
        HH: pad(date.getHours()), H: String(date.getHours()), mm: pad(date.getMinutes())
      };
      return format.replace(/yyyy|MM|M|dd|d|HH|H|mm/g, token => map[token]);
    }
  };
  c.getStaffName_ = () => 'S';
  c.resolveUserName_ = () => 'A';
  c.isAmbiguousUserName_ = () => false;
  c.isRegisteredStaffUser_ = () => true;
  c.sendVisitConfirmationPushFromLiff_ = () => {};
  c.refreshCouponAfterVisit_ = () => ({ warning: '' });
  return c;
}

test('LIFF cannot double-register an end visit already recorded via LINE for the same schedule', () => {
  const c = context();
  // Schedule row 2: staff S / user A, still shows "訪問中" because the LINE-channel
  // end message never touches the schedule sheet's status column.
  const scheduleSheet = makeSheet([
    ['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', 'eventId', '状態', '更新日時', '更新内容'],
    [new Date(2026, 8, 16), 'S', 'A', new Date(2026, 8, 16), '', '', '', '', '訪問中', '', '']
  ]);
  // Visit result sheet already has both a LIFF start and a LINE-received end for the same day.
  const resultSheet = makeSheet([
    ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻', '本文', 'LINEユーザーID'],
    [new Date(2026, 8, 16, 10, 0), '開始', 'S', 'A', '9/16', '10:00', 'LIFF実績登録：A 開始', 'id'],
    [new Date(2026, 8, 16, 10, 45), '終了', 'S', 'A', '9/16', '10:45', 'A 終了', 'line-id']
  ]);
  const otherSheets = {};
  const ss = {
    getSheetByName: name => {
      if (name === '訪問実績') return resultSheet;
      if (name === '訪問予定') return scheduleSheet;
      if (!otherSheets[name]) otherSheets[name] = makeSheet([]);
      return otherSheets[name];
    },
    insertSheet: name => { otherSheets[name] = makeSheet([]); return otherSheets[name]; }
  };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };

  const resultRowsBefore = resultSheet.rows.length;
  const scheduleBefore = scheduleSheet.rows[1].slice();

  const result = c.recordVisitFromLiff_('id', 'A', '終了', '2026-09-16', '11:00', '2');

  assert.equal(result.success, false);
  assert.equal(result.duplicate, true);
  assert.match(result.message, /別の経路で登録されています/);
  assert.equal(resultSheet.rows.length, resultRowsBefore, '訪問実績シートに新しい行を追加してはいけない');
  assert.deepEqual(scheduleSheet.rows[1], scheduleBefore, '予定シートの状態列を変更してはいけない');
});

test('a genuinely new end visit for a different schedule is still accepted', () => {
  const c = context();
  const scheduleSheet = makeSheet([
    ['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', 'eventId', '状態', '更新日時', '更新内容'],
    [new Date(2026, 8, 16), 'S', 'A', new Date(2026, 8, 16), '', '', '', '', '訪問中', '', '']
  ]);
  const resultSheet = makeSheet([
    ['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻', '本文', 'LINEユーザーID'],
    [new Date(2026, 8, 16, 10, 0), '開始', 'S', 'A', '9/16', '10:00', 'LIFF実績登録：A 開始', 'id']
  ]);
  const otherSheets = {};
  const ss = {
    getSheetByName: name => {
      if (name === '訪問実績') return resultSheet;
      if (name === '訪問予定') return scheduleSheet;
      if (!otherSheets[name]) otherSheets[name] = makeSheet([]);
      return otherSheets[name];
    },
    insertSheet: name => { otherSheets[name] = makeSheet([]); return otherSheets[name]; }
  };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };

  const result = c.recordVisitFromLiff_('id', 'A', '終了', '2026-09-16', '11:00', '2');

  assert.equal(result.success, true);
  assert.equal(resultSheet.rows.length, 3);
  assert.equal(scheduleSheet.rows[1][8], '完了');
});

test('LINE-message visit registration holds a script lock around the duplicate-check and append', () => {
  const code = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  const c = vm.createContext({ console: { error() {} } });
  vm.runInContext(code, c);
  c.Utilities = {
    formatDate: (date, _tz, format) => {
      const pad = n => String(n).padStart(2, '0');
      const map = { M: String(date.getMonth() + 1), d: String(date.getDate()), H: String(date.getHours()), mm: pad(date.getMinutes()) };
      return format.replace(/M|d|H|mm/g, token => map[token]);
    }
  };
  let locked = 0, released = 0, appendedWhileLocked = null;
  c.LockService = { getScriptLock: () => ({
    tryLock: () => { locked++; return true; },
    releaseLock: () => { released++; }
  }) };
  const resultSheet = makeSheet([['日時', '種別', 'スタッフ名', '利用者名', '実施日', '時刻', '本文', 'LINEユーザーID']]);
  const staffUserSheet = makeSheet([['スタッフ名', '利用者名'], ['S', 'A']]);
  const userMasterSheet = makeSheet([['利用者名'], ['A']]);
  const originalAppend = resultSheet.appendRow;
  resultSheet.appendRow = row => { appendedWhileLocked = locked > released; originalAppend(row); };
  const otherSheets = {};
  const ss = {
    getSheetByName: name => {
      if (name === '訪問実績') return resultSheet;
      if (name === 'スタッフ利用者マスタ') return staffUserSheet;
      if (name === '利用者マスタ') return userMasterSheet;
      if (name === '訪問予定') return makeSheet([]);
      if (!otherSheets[name]) otherSheets[name] = makeSheet([]);
      return otherSheets[name];
    },
    insertSheet: name => { otherSheets[name] = makeSheet([]); return otherSheets[name]; }
  };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };
  c.ContentService = { createTextOutput: () => ({ setMimeType() { return this; } }), MimeType: { TEXT: 'text' } };
  c.sendReplyMessages_ = () => {};
  c.getStaffName_ = () => 'S';
  c.getLineDisplayNameFromEvent_ = () => 'S';
  c.saveLineUserDirectory_ = () => {};
  c.isRegisteredStaffUser_ = () => true;
  c.refreshCouponAfterVisit_ = () => ({ warning: '' });

  c.doPost({ postData: { contents: JSON.stringify({
    events: [{ type: 'message', source: { userId: 'line-id' }, replyToken: 'rt',
      message: { type: 'text', text: 'A 終了' } }]
  }) } });

  assert.equal(locked, 1);
  assert.equal(released, 1);
  assert.equal(appendedWhileLocked, true, '実績の追記はロック保持中に行われるべき');
  assert.equal(resultSheet.rows.length, 2);
});
