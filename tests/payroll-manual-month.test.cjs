const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function setup(now) {
  const c = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  const RealDate = Date;
  c.Date = class extends RealDate {
    constructor(...args) { if (args.length === 0) super(now.getTime()); else super(...args); }
  };
  vm.runInContext('Date = this.Date', c);
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: () => null }), getUi: () => { throw new Error('no ui'); } };
  c.Utilities = { formatDate: d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') };
  const calls = [];
  c.createWageLedger = () => calls.push('ledger');
  c.createPayrollPdfFromTemplate = (_ss, ym) => calls.push('pdf:' + ym);
  c.createGmoTransferCsv = (_ss, ym) => calls.push('csv:' + ym);
  return { c, calls };
}

test('the manual payslip menu makes only the previous month, not the month in progress', () => {
  const s = setup(new Date(2026, 9, 1, 10));
  s.c.runPayrollFilesAfterReview();
  assert.deepEqual(s.calls, ['ledger', 'pdf:2026-09', 'csv:2026-09']);
});

test('in January the previous month is December of the year before', () => {
  const s = setup(new Date(2027, 0, 5, 10));
  s.c.runPayrollFilesAfterReview();
  assert.deepEqual(s.calls, ['ledger', 'pdf:2026-12', 'csv:2026-12']);
});
