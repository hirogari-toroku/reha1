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
    clearContents() { rows.length = 0; },
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

function formatDate(date, _tz, format) {
  const pad = n => String(n).padStart(2, '0');
  return format
    .replace('yyyy', date.getFullYear())
    .replace('MM', pad(date.getMonth() + 1))
    .replace('dd', pad(date.getDate()))
    .replace('HH', pad(date.getHours()))
    .replace('H', date.getHours())
    .replace('mm', pad(date.getMinutes()));
}

const STAFF_HEADERS = ['スタッフ名', '銀行コード', '支店番号', '預金種目', '口座番号', '受取人名', '振込先登録状況', 'スタッフフォルダID'];
const LEDGER_HEADERS = ['支払月日', '氏名', '性別', '労働日数', '労働時間数', '早出残業時間数', '深夜労働時間数', '基本賃金', '賃金所定時間外割増', '通勤手当', '奨励手当', '紹介手当', '運営報酬', '合計', '控除額', '実物給与', '備考'];
const ledgerRow = (payDate, name, payment) => [payDate, name, '', '', '', '', '', '', '', '', '', '', '', '', '', payment, ''];
const PAYROLL_HEADERS = ['年月', '振込日', 'スタッフ名', '訪問件数', '基本給', '交通費', '奨励手当', '紹介手当', '運営報酬', '課税対象額', '所得税', '支給額', '備考'];

// 2026年10月分を対象に、11月1日に動かす想定。
function setup({ schedules = [], visits = [], payroll = [], staff, ledger = [], assignments } = {}) {
  const sheets = {
    '訪問予定': makeSheet([['登録日時', 'スタッフ', '利用者', '訪問日', '', '', '', '', '状態'], ...schedules]),
    '訪問実績': makeSheet([['受付日時', '種別', 'スタッフ', '利用者', '訪問日', '時刻'], ...visits]),
    '給与集計': makeSheet([PAYROLL_HEADERS, ...payroll]),
    'スタッフマスタ': makeSheet([STAFF_HEADERS, ...(staff || [
      ['スタッフA', '0001', '001', '1', '1234567', 'ｴｰ', '登録済', 'folderA'],
      ['スタッフB', '0001', '001', '1', '7654321', 'ﾋﾞｰ', '登録済', 'folderB']
    ])]),
    'スタッフ利用者マスタ': makeSheet([['スタッフ', '利用者', 'LINE表示名', '単価'], ...(assignments || [
      ['スタッフA', '利用者X', '', 3000],
      ['スタッフA', '利用者Z', '', 3000],
      ['スタッフB', '利用者Y', '', 3000]
    ])]),
    '賃金台帳': makeSheet([LEDGER_HEADERS, LEDGER_HEADERS.map(() => ''), ...ledger])
  };
  const c = vm.createContext({ console, Date });
  for (const name of ['コード.js', 'PayrollAuto.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas', name), 'utf8'), c);
  }
  const ss = {
    getSheetByName: name => sheets[name] || null,
    insertSheet: name => { sheets[name] = makeSheet([]); return sheets[name]; }
  };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss, openById: () => ss, getUi: () => { throw new Error('no ui'); } };
  c.Utilities = { formatDate };
  const props = {};
  c.PropertiesService = { getScriptProperties: () => ({
    getProperty: k => (k in props ? props[k] : null),
    setProperty: (k, v) => { props[k] = v; }
  }) };
  c.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  const pushes = [];
  c.sendPushMessage_ = (_ss, to, _n, text) => { pushes.push({ to, text }); return { success: true }; };
  // 単価・交通費の区分判定は別のテストで検証済み。ここでは担当の有無だけを使う。
  c.getStaffUserMasterMap_ = () => Object.fromEntries(sheets['スタッフ利用者マスタ'].rows.slice(1)
    .map(r => [c.normalizeName_(r[0]) + '_' + c.normalizeName_(r[1]), { unitPay: r[3], travelCost: 0 }]));
  c.isAdminLiffUser_ = (_ss, id) => id === 'ADMIN';
  c.adminDeniedResponse_ = () => ({ success: false, message: 'denied' });
  c.getPayrollTransferDate_ = () => '2026/11/13';
  // 集計そのもの（既存の処理）は別に検証済み。ここでは用意した給与集計を使う。
  c.updateScheduleVisitComparisonCore_ = () => ({ success: true });
  c.createPayrollSummary = () => {};
  const made = [];
  c.createWageLedger = (_ss, ym) => made.push('ledger:' + ym);
  c.createPayrollPdfFromTemplate = (_ss, ym) => { made.push('pdf:' + ym); return { createdCount: 2, failedFiles: [], noStaffFolderFiles: [] }; };
  c.createGmoTransferCsv = (_ss, ym) => { made.push('csv:' + ym); return { success: true, createdCsvCount: 1, skippedCsvFiles: [], unregisteredStaffs: [] }; };
  c.movePayrollPdfsToStaffFoldersCore_ = () => { made.push('move'); return { movedCount: 2, skippedFiles: [], csvMovedCount: 1 }; };
  return { c, ss, sheets, pushes, made, props };
}

const d = (m, day, h = 10, min = 0) => new Date(2026, m - 1, day, h, min);

const cleanMonth = {
  schedules: [
    [d(9, 25), 'スタッフA', '利用者X', d(10, 3), '', '', '', '', '予定'],
    [d(9, 25), 'スタッフB', '利用者Y', d(10, 4), '', '', '', '', '予定']
  ],
  visits: [
    [d(10, 3, 10), '開始', 'スタッフA', '利用者X', d(10, 3), '10:00'],
    [d(10, 3, 11), '終了', 'スタッフA', '利用者X', d(10, 3), '11:00'],
    [d(10, 4, 10), '開始', 'スタッフB', '利用者Y', d(10, 4), '10:00'],
    [d(10, 4, 11), '終了', 'スタッフB', '利用者Y', d(10, 4), '11:00']
  ],
  payroll: [
    ['2026-09', '2026/10/15', 'スタッフA', 5, 15000, 1000, 0, 0, 0, 15000, 459, 15541, ''],
    ['2026-09', '2026/10/15', 'スタッフB', 4, 12000, 800, 0, 0, 0, 12000, 367, 12433, ''],
    ['2026-10', '2026/11/13', 'スタッフA', 6, 18000, 1200, 0, 0, 0, 18000, 551, 18649, ''],
    ['2026-10', '2026/11/13', 'スタッフB', 4, 12000, 800, 0, 0, 0, 12000, 367, 12433, '']
  ]
};

function withNow(c, date, fn) {
  const RealDate = Date;
  c.Date = class extends RealDate {
    constructor(...args) { if (args.length === 0) super(date.getTime()); else super(...args); }
  };
  vm.runInContext('Date = this.Date', c);
  try { return fn(); } finally { c.Date = RealDate; vm.runInContext('Date = this.Date', c); }
}

test('previous month rolls over the year', () => {
  const s = setup();
  assert.equal(s.c.previousMonthKey_(new Date(2027, 0, 1)), '2026-12');
  assert.equal(s.c.previousMonthKey_(new Date(2026, 10, 1)), '2026-10');
});

test('even a clean month stops before payslips and waits for the admin', () => {
  const s = setup(cleanMonth);
  const result = withNow(s.c, d(11, 1, 8, 0), () => s.c.runMonthlyPayrollAuto());
  assert.equal(result.status, '確認待ち');
  assert.deepEqual(s.made, [], '確認前に明細を作らない');
  assert.equal(s.pushes.length, 0, 'LINEは送らない');
  assert.equal(s.sheets['給与チェック'].rows[1][1], '問題なし');
});

test('approving the waiting month creates payslips, folders and the transfer CSV', () => {
  const s = setup(cleanMonth);
  withNow(s.c, d(11, 1, 8, 0), () => s.c.runMonthlyPayrollAuto());
  const result = withNow(s.c, d(11, 1, 9, 0), () => s.c.approvePayrollAndCreateFilesCore_(s.ss));
  assert.equal(result.status, '作成済');
  assert.deepEqual(s.made, ['ledger:2026-10', 'pdf:2026-10', 'csv:2026-10', 'move']);
  assert.match(result.message, /銀行への登録/);
  assert.equal(s.pushes.length, 0);
});

test('suspicious records stop before payslips and are listed for review', () => {
  const s = setup({
    schedules: [
      ...cleanMonth.schedules,
      [d(9, 25), 'スタッフA', '利用者X', d(10, 10), '', '', '', '', '予定'],   // 実績なし
      [d(9, 25), 'スタッフB', '利用者Y', d(10, 11), '', '', '', '', '予定'],   // 開始のみ
      [d(9, 25), 'スタッフB', '利用者Y', d(10, 31), '', '', '', '', '予定']    // 開始のみ
    ],
    visits: [
      ...cleanMonth.visits,
      [d(10, 12, 10), '開始', 'スタッフA', '利用者Z', d(10, 12), '10:00'],   // 予定外・短すぎ
      [d(10, 12, 10, 5), '終了', 'スタッフA', '利用者Z', d(10, 12), '10:05'],
      [d(10, 11, 10), '開始', 'スタッフB', '利用者Y', d(10, 11), '10:00'],
      [d(10, 31, 15), '開始', 'スタッフB', '利用者Y', d(10, 31), '15:00']
    ],
    payroll: [
      ...cleanMonth.payroll.slice(0, 3),
      ['2026-10', '2026/11/13', 'スタッフB', 4, 14000, 800, 0, 0, 0, 14000, 428, 14372, '利用者W']  // 単価変化・除外
    ],
    staff: [
      ['スタッフA', '0001', '001', '1', '1234567', 'ｴｰ', '登録済', 'folderA'],
      ['スタッフB', '0001', '001', '1', '7654321', 'ﾋﾞｰ', '要確認', 'folderB']
    ]
  });

  const result = withNow(s.c, d(11, 1, 8, 30), () => s.c.runMonthlyPayrollAuto());
  assert.equal(result.status, '要確認');
  assert.deepEqual(s.made, [], '要確認があるときは明細を作らない');

  const items = result.issues.filter(i => i.level === '要確認').map(i => i.item);
  for (const item of ['未実施', '開始のみ', '予定外', '訪問時間', '単価の変化', '集計から除外', '口座']) {
    assert.ok(items.includes(item), item + ' が検出されるべき: ' + items.join(','));
  }

  assert.equal(s.pushes.length, 0);
  assert.equal(s.sheets['給与チェック'].rows[1][1], '要確認');
});

test('approving after review creates files for the month that was held', () => {
  const s = setup(cleanMonth);
  s.c.savePayrollAutoState_('2026-10', '要確認');
  const result = withNow(s.c, d(12, 3, 9), () => s.c.approvePayrollAndCreateFilesCore_(s.ss));
  assert.equal(result.ym, '2026-10');
  assert.deepEqual(s.made, ['ledger:2026-10', 'pdf:2026-10', 'csv:2026-10', 'move']);
  assert.equal(JSON.parse(s.props.PAYROLL_AUTO_STATE).status, '作成済');
});

test('approval goes ahead even when review items remain (the admin has checked them)', () => {
  const s = setup({ ...cleanMonth, staff: [['スタッフA', '0001', '001', '1', '1234567', 'ｴｰ', '要確認', 'folderA'], ['スタッフB', '0001', '001', '1', '7654321', 'ﾋﾞｰ', '登録済', 'folderB']] });
  s.c.savePayrollAutoState_('2026-10', '要確認');
  withNow(s.c, d(11, 3, 9), () => s.c.approvePayrollAndCreateFilesCore_(s.ss));
  assert.ok(s.made.includes('move'));
});

test('recheck only re-runs the checks and never creates files', () => {
  const s = setup(cleanMonth);
  const result = withNow(s.c, d(11, 2, 9), () => s.c.recheckPayrollCore_(s.ss));
  assert.match(result.message, /要確認 0件/);
  assert.deepEqual(s.made, []);
});

test('a summary error stops the run before any payslip is made', () => {
  const s = setup(cleanMonth);
  s.c.createPayrollSummary = () => { throw new Error('契約済みの給与単価を入力してください。'); };
  const result = withNow(s.c, d(11, 1, 8, 30), () => s.c.runMonthlyPayrollAuto());
  assert.equal(result.status, 'エラー');
  assert.deepEqual(s.made, []);
  assert.equal(s.pushes.length, 0);
});

test('admin-only LIFF actions refuse other users', () => {
  const s = setup(cleanMonth);
  assert.equal(s.c.adminPayrollApproveFromLiff_('someone').success, false);
  assert.deepEqual(s.made, []);
});

test('a message logged as a visit under a non-client name is noted but does not stop payroll', () => {
  const s = setup({ ...cleanMonth, visits: [...cleanMonth.visits, [d(10, 5, 9), '開始', 'スタッフA', 'お疲れ', d(10, 5), '9:00']] });
  const result = withNow(s.c, d(11, 1, 8), () => s.c.recheckPayrollCore_(s.ss));
  const issue = result.issues.find(i => i.item === '利用者以外');
  assert.equal(issue.level, '参考');
  assert.equal(result.issues.filter(i => i.level === '要確認').length, 0);
});

test('a surname-only visit matches the full-name schedule instead of showing as missed and unplanned', () => {
  const s = setup({
    schedules: [[d(9, 25), 'スタッフA', '利用者X', d(10, 3), '', '', '', '', '予定']],
    visits: [
      [d(10, 3, 10), '開始', 'スタッフA', '利用者', d(10, 3), '10:00'],
      [d(10, 3, 11), '終了', 'スタッフA', '利用者', d(10, 3), '11:00']
    ],
    assignments: [['スタッフA', '利用者X', '', 3000]],
    payroll: cleanMonth.payroll.filter(r => r[2] === 'スタッフA')
  });
  const result = withNow(s.c, d(11, 1, 8), () => s.c.recheckPayrollCore_(s.ss));
  const items = result.issues.map(i => i.item);
  assert.ok(!items.includes('未実施') && !items.includes('予定外'), items.join(','));
});

test('an end registered twice within minutes is a note, ends far apart need review', () => {
  const twice = (endA, endB) => setup({
    ...cleanMonth,
    visits: [...cleanMonth.visits, [d(10, 3, 11, 9), '終了', 'スタッフA', '利用者X', d(10, 3), endB]],
    payroll: cleanMonth.payroll
  });
  const near = twice('11:00', '11:09');
  const nearIssue = withNow(near.c, d(11, 1, 8), () => near.c.recheckPayrollCore_(near.ss)).issues.find(i => i.item === '二重登録');
  assert.equal(nearIssue.level, '参考');
  const far = twice('11:00', '14:00');
  const farIssue = withNow(far.c, d(11, 1, 8), () => far.c.recheckPayrollCore_(far.ss)).issues.find(i => i.item === '重複');
  assert.equal(farIssue.level, '要確認');
});

test('a paid month whose recalculated pay no longer matches the wage ledger needs review', () => {
  const s = setup({
    ...cleanMonth,
    ledger: [ledgerRow('2026/10/15', 'スタッフA', 15541), ledgerRow('2026/10/15', 'スタッフB', 10000)]
  });
  const result = withNow(s.c, d(11, 1, 8), () => s.c.recheckPayrollCore_(s.ss));
  const changed = result.issues.filter(i => i.item === '支払済み月の変化');
  assert.equal(changed.length, 1);
  assert.equal(changed[0].staffName, 'スタッフB');
  assert.match(changed[0].detail, /10000円 → 12433円/);
});

test('a payroll memo that only lists same-day double registrations is a note', () => {
  const payroll = cleanMonth.payroll.map(r => r[0] === '2026-10' && r[2] === 'スタッフB' ? [...r.slice(0, 12), '利用者Y（2026-10-04 重複除外）'] : r);
  const s = setup({ ...cleanMonth, payroll });
  const result = withNow(s.c, d(11, 1, 8), () => s.c.recheckPayrollCore_(s.ss));
  assert.equal(result.issues.find(i => i.item === '集計から除外').level, '参考');
});
