const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function makeSheet(rows) {
  return {
    rows,
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
    setFrozenRows() {},
    getDataRange() { return this.getRange(1, 1, rows.length, this.getLastColumn()); },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = (rows[row - 1 + i] || [])[col - 1 + j];
          return v === undefined ? '' : v;
        })),
      setValues: values => values.forEach((line, i) => {
        while (rows.length < row + i) rows.push([]);
        line.forEach((v, j) => { rows[row - 1 + i][col - 1 + j] = v; });
      })
    })
  };
}

function setup(sheets) {
  const c = vm.createContext({ console });
  for (const name of ['コード.js', 'Maintenance.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas', name), 'utf8'), c);
  }
  const store = {};
  const ss = { getSheetByName: name => sheets[name] || null, insertSheet: name => { sheets[name] = makeSheet([]); return sheets[name]; } };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss, openById: () => ss, flush() {} };
  c.CacheService = { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; } }) };
  c.Utilities = {
    formatDate: (date, _tz, format) => {
      const pad = n => String(n).padStart(2, '0');
      if (format === 'yyyy-MM') return date.getFullYear() + '-' + pad(date.getMonth() + 1);
      if (format === 'yyyy-MM-dd') return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate());
      return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) + '_' + pad(date.getHours()) + pad(date.getMinutes());
    }
  };
  const pushes = [];
  c.sendPushMessage_ = (_ss, to, _n, text) => { pushes.push({ to, text }); return { success: true }; };
  c.saveLiffOperationLog_ = () => {};
  c.isAdminLiffUser_ = (_ss, id) => id === 'ADMIN';
  c.adminDeniedResponse_ = () => ({ success: false, message: 'denied' });
  return { c, ss, sheets, pushes, store };
}

const thisMonth = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), 5, 10, 0); };
const lastMonth = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth() - 1, 5, 10, 0); };

function logRows(sentThis, sentLast, failed) {
  const rows = [['日時', '区分', '相手']];
  for (let i = 0; i < sentThis; i++) rows.push([thisMonth(), '送信', 'x']);
  for (let i = 0; i < sentLast; i++) rows.push([lastMonth(), '送信', 'x']);
  for (let i = 0; i < failed; i++) rows.push([thisMonth(), '送信失敗', 'x']);
  rows.push([thisMonth(), '受信', 'x']);
  return rows;
}

test('only pushes are counted against the free allowance; replies and receipts are not', () => {
  const s = setup({ 'LINE送受信ログ': makeSheet(logRows(12, 30, 2)) });
  const usage = s.c.getLinePushUsage_(s.ss);
  assert.equal(usage.sent, 12);
  assert.equal(usage.failed, 2);
  assert.equal(usage.lastMonthSent, 30);
  assert.equal(usage.limit, 200);
  assert.equal(usage.remaining, 188);
  assert.equal(usage.warning, false);
  assert.equal(usage.overLimit, false);
});

test('the allowance warning and over-limit flags turn on at 150 and 200', () => {
  const below = setup({ 'LINE送受信ログ': makeSheet(logRows(149, 0, 0)) });
  assert.equal(below.c.getLinePushUsage_(below.ss).warning, false, '149 is still fine');
  const warn = setup({ 'LINE送受信ログ': makeSheet(logRows(150, 0, 0)) });
  assert.equal(warn.c.getLinePushUsage_(warn.ss).warning, true);
  const over = setup({ 'LINE送受信ログ': makeSheet(logRows(200, 0, 0)) });
  assert.equal(over.c.getLinePushUsage_(over.ss).overLimit, true);
  assert.equal(over.c.getLinePushUsage_(over.ss).remaining, 0);
});

test('the monthly check counts visits, schedules and deposits, and flags problems', () => {
  const visitRows = [['日時', '種別', 'スタッフ', '利用者', '実施日']];
  for (let i = 0; i < 3; i++) visitRows.push([thisMonth(), '開始', 'S', 'U', thisMonth()]);
  for (let i = 0; i < 5; i++) visitRows.push([thisMonth(), '終了', 'S', 'U', thisMonth()]);
  const s = setup({
    '訪問実績': makeSheet(visitRows),
    '訪問予定': makeSheet([['登録', 'スタッフ', '利用者', '訪問日', '', '', '', '', '状態'],
      [thisMonth(), 'S', 'U', thisMonth(), '', '', '', '', '完了'],
      [thisMonth(), 'S', 'U', thisMonth(), '', '', '', '', '予定'],
      [thisMonth(), 'S', 'U', thisMonth(), '', '', '', '', 'キャンセル']]),
    '入出金明細': makeSheet([['日付', '摘要', '入金'], [thisMonth(), '振込', 30000], [thisMonth(), '振込', 60000], [lastMonth(), '振込', 10000]]),
    '回数券管理': makeSheet([['利用者名', '回数券残数'], ['A', -1], ['B', 3]]),
    'LINE送受信ログ': makeSheet(logRows(160, 0, 1))
  });
  const check = s.c.buildMonthlyCheck_(s.ss, s.c.Utilities.formatDate(thisMonth(), 'x', 'yyyy-MM'));
  assert.equal(check.startCount, 3);
  assert.equal(check.endCount, 5);
  assert.equal(check.scheduleCount, 2, 'cancelled schedules are excluded');
  assert.equal(check.completedCount, 1);
  assert.equal(check.depositCount, 2);
  assert.equal(check.depositTotal, 90000);
  assert.equal(check.couponNegative, 1);
  assert.deepEqual(check.issues.length, 4, check.issues.join(' / '));
  assert.match(check.issues.join(' / '), /終了実績が開始実績より多い/);
  assert.match(check.issues.join(' / '), /回数券残数がマイナス/);
  assert.match(check.issues.join(' / '), /無料枠に近い/);
  assert.match(check.issues.join(' / '), /送信失敗/);
});

test('a clean month reports no problems and writes one row per checked month', () => {
  const s = setup({
    '訪問実績': makeSheet([['日時', '種別', 'スタッフ', '利用者', '実施日'], [thisMonth(), '開始', 'S', 'U', thisMonth()], [thisMonth(), '終了', 'S', 'U', thisMonth()]]),
    '訪問予定': makeSheet([['登録', 'スタッフ', '利用者', '訪問日', '', '', '', '', '状態'], [thisMonth(), 'S', 'U', thisMonth(), '', '', '', '', '完了']]),
    '入出金明細': makeSheet([['日付', '摘要', '入金'], [thisMonth(), '振込', 30000]]),
    '回数券管理': makeSheet([['利用者名', '回数券残数'], ['A', 2]]),
    'LINE送受信ログ': makeSheet(logRows(5, 5, 0))
  });
  const result = s.c.runMonthlyMaintenanceCheck();
  assert.equal(result.checks[0].issues.length, 0);
  const sheet = s.sheets['月次点検'];
  assert.ok(sheet, 'the check sheet is created');
  assert.equal(sheet.rows.length, 1 + result.checks.length);
  assert.equal(s.pushes.length, 0, 'nothing to report, so no LINE message');
});

test('problems notify the admin once a day', () => {
  const s = setup({
    '訪問実績': makeSheet([['日時', '種別', 'スタッフ', '利用者', '実施日'], [thisMonth(), '終了', 'S', 'U', thisMonth()]]),
    '訪問予定': makeSheet([['登録', 'スタッフ', '利用者', '訪問日']]),
    '入出金明細': makeSheet([['日付', '摘要', '入金']]),
    '回数券管理': makeSheet([['利用者名', '回数券残数']]),
    'LINE送受信ログ': makeSheet(logRows(0, 0, 0))
  });
  s.c.runMonthlyMaintenanceCheck();
  assert.equal(s.pushes.length, 1);
  assert.match(s.pushes[0].text, /月次点検/);
  s.c.runMonthlyMaintenanceCheck();
  assert.equal(s.pushes.length, 1, 'the same warning is not sent twice in a day');
});

test('backup keeps only the newest generations and is set up by a time-driven trigger', () => {
  const src = fs.readFileSync(path.join(__dirname, '../gas/Maintenance.js'), 'utf8');
  assert.match(src, /ScriptApp\.newTrigger\("runDailyBackup"\)\.timeBased\(\)/);
  assert.match(src, /ScriptApp\.newTrigger\("runMonthlyMaintenanceCheck"\)\.timeBased\(\)/);

  const s = setup({});
  const trashed = [];
  const files = [];
  for (let i = 0; i < 10; i++) {
    files.push({ name: '自費リハ管理_2026-09-0' + i, created: new Date(2026, 8, i + 1), trashed: false });
  }
  files.push({ name: '別のファイル', created: new Date(2026, 8, 20), trashed: false });
  let index = 0;
  const iterator = {
    hasNext: () => index < files.length,
    next: () => {
      const f = files[index++];
      return {
        getName: () => f.name,
        getDateCreated: () => f.created,
        setTrashed: () => { f.trashed = true; trashed.push(f.name); }
      };
    }
  };
  s.c.MimeType = { GOOGLE_SHEETS: 'sheets' };
  const removed = s.c.removeOldBackups_({ getFilesByType: () => iterator });
  assert.equal(removed, 3, '10 backups minus 7 kept');
  assert.equal(trashed.length, 3);
  assert.deepEqual(trashed.sort(), ['自費リハ管理_2026-09-00', '自費リハ管理_2026-09-01', '自費リハ管理_2026-09-02']);
  assert.equal(files[files.length - 1].trashed, false, 'unrelated files are left alone');
});

test('the admin actions are wired and refuse non-admins', () => {
  const code = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  ['adminLineUsage', 'adminRunMonthlyCheck', 'adminSetupDailyBackup'].forEach(action => {
    assert.match(code, new RegExp('action === "' + action + '"'));
  });
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  ['adminLineUsage', 'adminRunMonthlyCheck', 'adminSetupDailyBackup'].forEach(action => {
    assert.match(html, new RegExp("runAdminAction\\('" + action + "'"));
  });
  const s = setup({ 'LINE送受信ログ': makeSheet(logRows(1, 0, 0)) });
  assert.equal(s.c.adminLineUsageFromLiff_('someone').success, false);
  assert.equal(s.c.adminLineUsageFromLiff_('ADMIN').success, true);
});

test('the check runs at midday and the backup at night', () => {
  const src = fs.readFileSync(path.join(__dirname, '../gas/Maintenance.js'), 'utf8');
  assert.match(src, /const BACKUP_TRIGGER_HOUR = 3;/);
  assert.match(src, /const MONTHLY_CHECK_TRIGGER_HOUR = 13;/);
  assert.match(src, /newTrigger\("runDailyBackup"\)\.timeBased\(\)\.atHour\(BACKUP_TRIGGER_HOUR\)/);
  assert.match(src, /newTrigger\("runMonthlyMaintenanceCheck"\)\.timeBased\(\)\.atHour\(MONTHLY_CHECK_TRIGGER_HOUR\)/);
});

test('the setup records the hours it used, and the status action compares them with the code', () => {
  const s = setup({});
  const created = [];
  const store = {};
  s.c.ScriptApp = {
    getProjectTriggers: () => created.map(name => ({ getHandlerFunction: () => name })),
    newTrigger: name => ({
      timeBased: () => ({ atHour: () => ({ everyDays: () => ({ create: () => created.push(name) }) }) })
    }),
    deleteTrigger: () => {}
  };
  s.c.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: k => (k in store ? store[k] : null),
      setProperty: (k, v) => { store[k] = v; }
    })
  };

  assert.match(s.c.adminMaintenanceStatusFromLiff_('ADMIN').message, /未設定/);
  s.c.setupDailyBackupTriggerCore_();
  const saved = JSON.parse(store.MAINTENANCE_TRIGGER_STATE);
  assert.equal(saved.backupHour, 3);
  assert.equal(saved.checkHour, 13);

  const status = s.c.adminMaintenanceStatusFromLiff_('ADMIN');
  assert.equal(status.maintenance.backup, true);
  assert.equal(status.maintenance.check, true);
  assert.match(status.message, /設定済みです/);
  assert.match(status.message, /13時台/);

  // 古い時刻で登録されていた場合は、押し直しを促す
  store.MAINTENANCE_TRIGGER_STATE = JSON.stringify({ backupHour: 3, checkHour: 7, setAt: '2026-09-25 09:00' });
  assert.match(s.c.adminMaintenanceStatusFromLiff_('ADMIN').message, /もう一度押してください/);
  assert.equal(s.c.adminMaintenanceStatusFromLiff_('someone').success, false);
});
