const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const ctx = vm.createContext({ console });
  for (const name of ['コード.js', '料金区分.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas', name), 'utf8'), ctx);
  }
  return ctx;
}
function sheet(rows) {
  return { getDataRange: () => ({ getValues: () => rows }) };
}
test('legacy distance calculation and custom coupon prices are unchanged', () => {
  const c = context();
  assert.equal(c.pricingTravel_('', 25, 123).payroll, 200);
  assert.equal(c.pricingTravel_('旧料金', 25, 123).direct, 800);
  assert.equal(c.pricingTravel_('', '', 123).payroll, 123);
  const users = c.getCouponUsers_(sheet([['利用者名', '', '', '', '回数券単価'], ['A', '', '', '', 8800], ['B']]));
  assert.equal(users[0].unitPrice, 8800);
  assert.equal(users[1].unitPrice, 7500);
});
test('new travel has no 10km cap and no user charge', () => {
  const c = context();
  for (const [distance, amount] of [[0, 0], [9, 180], [10, 200], [25, 500], [12.5, 250]]) {
    const result = c.pricingTravel_('新料金', distance, 999);
    assert.equal(result.payroll, amount);
    assert.equal(result.direct, 0);
  }
  for (const bad of ['', undefined, -1, 'bad']) assert.throws(() => c.pricingTravel_('新料金', bad, 99));
  assert.throws(() => c.pricingTravel_('要確認', 25, 99));
  assert.throws(() => c.pricingPolicy_('typo'));
});
test('new coupon price is 9250 independent of legacy E-column value', () => {
  const c = context();
  const users = c.getCouponUsers_(sheet([['利用者名', '', '', '', '回数券単価', '利用料金区分'], ['A', '', '', '', 7500, '新料金']]));
  assert.equal(users[0].unitPrice, 9250);
  c.validateNewCouponPayments_(users, { A: [{ amount: 37000 }, { amount: 74000 }] });
  assert.throws(() => c.validateNewCouponPayments_(users, { A: [{ amount: 11000 }] }));
  assert.throws(() => c.validateNewCouponPayments_(users, { A: [{ amount: 9250 }] }));
});
test('pending payments are blocked, old payment rules retained', () => {
  const c = context();
  assert.throws(() => c.validateNewCouponPayments_([{ userName: 'A', pricingPolicy: '要確認' }], { A: [{ amount: 37000 }] }));
  c.validateNewCouponPayments_([{ userName: 'A', pricingPolicy: '旧料金' }], { A: [{ amount: 30000 }] });
});
test('payroll uses independent relationship travel and staff base pay', () => {
  const c = context();
  const relation = sheet([
    ['スタッフ名', '利用者名', '', '給与単価', '給与交通費', '距離km', '', '', '交通費区分'],
    ['S', 'Old', '', 4500, 200, 25, 800, '', ''],
    ['S', 'New', '', 4500, 200, 25, 800, '', '新料金'],
    ['N', 'New', '', '', '', 25, '', '', '新料金']
  ]);
  const staff = sheet([['スタッフ名', '報酬区分'], ['S', ''], ['N', '新料金']]);
  const map = c.getStaffUserMasterMap_({ getSheetByName: name => name === 'スタッフマスタ' ? staff : relation });
  assert.equal(map.S_Old.travelCost, 200);
  assert.equal(map.S_New.travelCost, 500);
  assert.equal(map.S_New.unitPay, 4500);
  assert.match(map.N_New.pricingError, /給与単価/);
});
test('travel refresh leaves every legacy cell untouched', () => {
  const c = context();
  const writes = [];
  const rows = [
    ['スタッフ名', '利用者名', '', '', '', '', '', '', '交通費区分'],
    ['S', 'Old', '', 4500, 777, 25, 888, '例外契約', ''],
    ['S', 'New', '', 4500, 0, 25, 0, '', '新料金']
  ];
  const data = { ...sheet(rows), getRange: (...args) => ({ setValues: values => writes.push([args, values]) }) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: () => data }) };
  c.updateTravelCostsInStaffUserMaster();
  assert.equal(writes.length, 1);
  assert.equal(writes[0][0][0], 3);
  assert.equal(writes[0][1][0][0], 500);
  assert.equal(writes[0][1][0][2], 0);
});
test('travel validation fails before writing any new row', () => {
  const c = context();
  let writes = 0;
  const data = { ...sheet([
    ['S', 'U', '', '', '', '', '', '', '交通費区分'],
    ['S', 'New1', '', 4500, 0, 25, 0, '', '新料金'],
    ['S', 'New2', '', 4500, 0, '', 0, '', '新料金']
  ]), getRange: () => ({ setValues: () => writes++ }) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: () => data }) };
  assert.throws(() => c.updateTravelCostsInStaffUserMaster());
  assert.equal(writes, 0);
});
test('setup endpoint requires existing admin authorization', () => {
  const c = context();
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ({}) };
  c.isAdminLiffUser_ = () => false;
  c.adminDeniedResponse_ = () => ({ denied: true });
  c.setupPricingPolicyColumnsCore_ = () => { throw new Error('must not run'); };
  assert.equal(c.adminSetupPricingPolicy_('outsider').denied, true);
});

test('September raise preserves August and does not schedule a November raise', () => {
  const c = context();
  const rows = [['スタッフ名', '利用者名', '適用開始月', '給与単価'],
    ['S', 'U', '2026-09', 3500], ['S', 'U', '1900-01', 2500]];
  const history = c.getPayrollUnitPayHistory_({ getSheetByName: () => sheet(rows) });
  for (const month of ['2026-05', '2026-08']) assert.equal(c.payrollUnitPayForMonth_(history, 'S_U', month, 3500), 2500);
  for (const month of ['2026-09', '2026-10', '2026-11', '2027-09']) assert.equal(c.payrollUnitPayForMonth_(history, 'S_U', month, 4500), 3500);
  assert.equal(c.payrollUnitPayForMonth_(history, 'Other_U', '2026-11', 3000), 3000);
});
test('history errors do not silently fall back to current pay', () => {
  const c = context();
  const headers = ['スタッフ名', '利用者名', '適用開始月', '給与単価'];
  for (const rows of [
    [['S', 'U', '2026-13', 3500]],
    [['S', 'U', '2026-09', 3500], ['S', 'U', '2026-09', 4000]],
    [['S', 'U', '2026-09', -100]],
    [['S', '', '2026-09', 3500]]
  ]) assert.throws(() => c.getPayrollUnitPayHistory_({ getSheetByName: () => sheet([headers, ...rows]) }));
  assert.throws(() => c.payrollUnitPayForMonth_({ S_U: [{ month: '2026-09', unitPay: 3500 }] }, 'S_U', '2026-08', 3500));
});
