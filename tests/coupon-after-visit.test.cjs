const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const c = vm.createContext({ console: { error() {} } });
  for (const file of ['コード.js', '料金区分.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas', file), 'utf8'), c);
  }
  return c;
}

test('only real-user end visits trigger recalculation', () => {
  const c = context();
  const calls = [];
  c.updateCouponManagement = name => calls.push(name);
  c.refreshCouponAfterVisit_({}, 'A', '開始');
  c.refreshCouponAfterVisit_({}, 'テスト利用者', '終了');
  c.refreshCouponAfterVisit_({}, 'A', '終了');
  assert.deepEqual(calls, ['A']);
});

test('update and logging failures never invalidate saved visits', () => {
  const c = context();
  c.updateCouponManagement = () => { throw Error('failed'); };
  c.saveLiffOperationLog_ = () => { throw Error('log failed'); };
  const result = c.refreshCouponAfterVisit_({}, 'A', '終了');
  assert.equal(result.success, false);
  assert.match(result.warning, /実績は保存済み/);
});

test('manual and targeted recalculation share a lock and release on errors', () => {
  const c = context();
  let releases = 0;
  c.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => releases++ }) };
  c.updateCouponManagementLocked_ = () => { throw Error('bad schema'); };
  assert.throws(() => c.updateCouponManagement('A'), /bad schema/);
  assert.throws(() => c.updateCouponManagement(), /bad schema/);
  assert.equal(releases, 2);
});

function fixture() {
  const c = context();
  const headers = ['利用者名', '回数券残数', '未払い残高', '最終入金日', '最終入金額', '1回単価', '不足金額', '基準日', '基準残数', '基準後入金回数', '基準後利用回数', '当月入金回数', '当月利用回数', '当月不足回数', '確認ステータス', 'メモ'];
  const writes = [];
  const coupon = { getLastRow: () => 3, getRange: (row, col, count, width) => ({
    getValues: () => row === 1 ? [headers] : [['A'], ['B']],
    setValues: values => writes.push({ row, col, count, width, values })
  }) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: name => name === '回数券管理' ? coupon : {} }), flush() {} };
  c.Utilities = { formatDate: () => '2026-09' };
  c.getExistingCouponMemoMap_ = () => ({ A: { memo: 'keep memo' } });
  c.getCouponUsers_ = () => [{ userName: 'A', unitPrice: 7500, pricingPolicy: '旧料金' }, { userName: 'B', unitPrice: 9250, pricingPolicy: '要確認' }];
  c.getCouponBaselineMap_ = () => ({ A: { baselineDate: '2026-08-01', baselineBalance: 4 } });
  c.getCouponPaymentEventsByUser_ = () => ({ B: [{ amount: 11000 }] });
  c.getUserPaymentMap_ = () => ({});
  c.getUserUsageMap_ = () => ({ A: { total: 1, monthly: { '2026-09': 1 } } });
  return { c, writes, headers };
}

test('repeated updates write only the target row and never subtract twice', () => {
  const { c, writes } = fixture();
  c.updateCouponManagementLocked_('A');
  c.updateCouponManagementLocked_('A');
  assert.equal(writes.length, 2);
  for (const write of writes) {
    assert.equal(write.row, 2);
    assert.equal(write.count, 1);
    assert.equal(write.values[0][1], 3);
    assert.equal(write.values[0][15], 'keep memo');
  }
});

test('unsafe pricing and changed headers prevent any target write', () => {
  const { c, writes, headers } = fixture();
  assert.throws(() => c.updateCouponManagementLocked_('B'), /料金区分/);
  headers[1] = 'unexpected';
  assert.throws(() => c.updateCouponManagementLocked_('A'), /列構成/);
  assert.equal(writes.length, 0);
});

test('duplicate end records on the same received date count once in existing usage rules', () => {
  const c = context();
  c.isOnOrAfterCouponStart_ = () => true;
  c.isAfterCouponBaseline_ = () => true;
  c.getCouponDateKey_ = value => value;
  c.getCouponMonthKey_ = () => '2026-09';
  const sheet = { getDataRange: () => ({ getValues: () => [[], ['2026-09-16', '終了', 'S', 'A'], ['2026-09-16', '終了', 'S', 'A']] }) };
  assert.equal(c.getUserUsageMap_(sheet, {}).A.total, 1);
});

test('LIFF saves the end visit before coupon refresh and returns success with warning', () => {
  const c = context();
  const events = [];
  const sheet = { appendRow: () => events.push('saved') };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }) };
  c.getStaffName_ = () => 'S';
  c.resolveUserName_ = () => 'A';
  c.normalizeLiffDateText_ = () => '2026-09-16';
  c.normalizeLiffTimeText_ = () => '12:00';
  c.isRegisteredStaffUser_ = () => true;
  c.isRecentDuplicateVisit_ = () => false;
  c.saveLineMessageLog_ = () => {};
  c.saveLiffOperationLog_ = () => {};
  c.sendVisitConfirmationPushFromLiff_ = () => {};
  c.updateCouponManagement = () => { events.push('coupon'); throw Error('unavailable'); };
  const result = c.recordVisitFromLiff_('id', 'A', '終了', '', '', '');
  assert.equal(result.success, true);
  assert.match(result.message, /回数券残数の更新は未完了/);
  assert.deepEqual(events, ['saved', 'coupon']);
});
