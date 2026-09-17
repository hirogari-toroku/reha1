const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), vm = require('node:vm'), path = require('node:path');
function fixture(json) {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.PropertiesService = {getScriptProperties: () => ({getProperty: k => k === 'LIFF_DISPLAY_MASTER_JSON' ? json : null})};
  c.CacheService = {getScriptCache: () => ({get: () => null, put() {}})};
  c.getLiffDisplayMasterCacheVersion_ = () => 'test';
  return c;
}
test('known property-map miss avoids display sheet reads for user entrance', () => {
  const c = fixture('{"staff":{"staffName":"staff"}}');
  c.SpreadsheetApp = {getActiveSpreadsheet() {throw Error('unnecessary sheet read');}};
  assert.equal(c.getLiffInitDataFromDisplayMaster_('user'), null);
  assert.equal(c.getLiffInitDataFromDisplayMaster_('staff').staffName, 'staff');
});
test('missing or invalid property maps retain sheet fallback', () => {
  for (const json of [null, '{', 'null', '[]']) {
    const c = fixture(json); let reads = 0;
    c.SpreadsheetApp = {getActiveSpreadsheet: () => ({getSheetByName() {reads++; return null;}})};
    assert.equal(c.getLiffInitDataFromDisplayMaster_('user'), null);
    assert.equal(reads, 1);
  }
});
test('coupon display uses one values read without rereading headers', () => {
  const c = fixture('{}'); let reads = 0;
  const ss = {getSheetByName: () => ({getLastRow: () => 2, getDataRange: () => ({getValues() {
    reads++; return [['利用者名','回数券残数','確認ステータス'],['テスト',2,'残数あり']];
  }})})};
  assert.equal(c.getCouponDisplayMap_(ss)['テスト'].balance, 2);
  assert.equal(reads, 1);
});
