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
    appendRow(row) { rows.push(row); },
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
      }),
      setFontWeight() { return this; }
    })
  };
}

const pad = n => String(n).padStart(2, '0');
function setup(sheets) {
  const c = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  const ss = { getSheetByName: name => sheets[name] || null };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss, getUi: () => { throw new Error('no ui'); } };
  c.Utilities = { formatDate: (date, _tz, format) => format === 'yyyy-MM'
    ? date.getFullYear() + '-' + pad(date.getMonth() + 1)
    : date.getFullYear() + '/' + pad(date.getMonth() + 1) + '/' + pad(date.getDate()) };
  // 単価・手当・祝日の判定は別のテストで検証済み。ここでは月の振り分けだけを見る。
  c.getStaffUserMasterMap_ = () => ({ [c.normalizeName_('スタッフA') + '_' + c.normalizeName_('利用者X')]: { unitPay: 3000, travelCost: 100 } });
  c.getPayrollUnitPayHistory_ = () => ({});
  c.getAnnualRaisePolicies_ = () => ({});
  c.getPayrollAllowanceMap_ = () => ({});
  c.getStaffFormalNameMap_ = () => ({});
  c.resolveUserName_ = (_ss, _staff, name) => name;
  c.payrollUnitPayForMonth_ = (_h, _k, _ym, pay) => pay;
  c.annualRaiseUnitPay_ = (_p, _k, _ym, pay) => pay;
  c.getPayrollTransferDate_ = ym => ym + '-15';
  return { c, ss };
}

test('a visit at the end of the month registered the next day counts for the month it happened', () => {
  const visits = makeSheet([
    ['受信日時', '種別', 'スタッフ名', '利用者名', '日付', '時刻', '原文', 'LINEユーザーID', '日時', '実施日'],
    [new Date(2026, 9, 3, 11), '終了', 'スタッフA', '利用者X', '10/3', '11:00'],
    [new Date(2026, 10, 1, 7), '終了', 'スタッフA', '利用者X', '10/31', '20:00'],
    [new Date(2026, 10, 2, 11), '終了', 'スタッフA', '利用者X', '11/2', '11:00']
  ]);
  const payroll = makeSheet([]);
  const s = setup({ '訪問実績': visits, '給与集計': payroll });
  s.c.createPayrollSummary(s.ss);
  const counts = Object.fromEntries(payroll.rows.slice(1).map(r => [r[0], r[3]]));
  assert.deepEqual(counts, { '2026-10': 2, '2026-11': 1 });
});

test('the explicit visit date column wins over the registration time', () => {
  const s = setup({});
  const date = s.c.getPayrollVisitDate_(['', '終了', 'A', 'X', '', '', '', '', '', '2026-08-31'], new Date(2026, 8, 1, 9));
  assert.equal(date.getMonth(), 7);
  assert.equal(date.getDate(), 31);
});

test('the wage ledger only writes the target month and keeps paid months as they were', () => {
  const payroll = makeSheet([
    ['年月', '振込日', 'スタッフ名', '訪問件数', '基本給', '交通費', '奨励手当', '紹介手当', '運営報酬', '課税対象額', '所得税', '支給額', '備考'],
    ['2026-09', '2026/10/15', 'スタッフA', 4, 12000, 400, 0, 0, 0, 12000, 367, 99999, ''],
    ['2026-10', '2026/11/13', 'スタッフA', 5, 15000, 500, 0, 0, 0, 15000, 459, 15041, '']
  ]);
  const ledger = makeSheet([[], [], ['2026/10/15', 'スタッフA', '', 4, 4, 0, 0, 12000, 0, 400, 0, 0, 0, 12400, 367, 12033, '']]);
  const s = setup({ '給与集計': payroll, '賃金台帳': ledger });
  s.c.createWageLedger(s.ss, '2026-10');
  assert.equal(ledger.rows[2][15], 12033, '支払済みの9月分は上書きしない');
  assert.equal(ledger.rows[3][0], '2026/11/13');
  assert.equal(ledger.rows[3][15], 15041);
});
