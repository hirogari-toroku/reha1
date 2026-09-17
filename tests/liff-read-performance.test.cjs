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

test('known staff initialization and important-info viewing do not write directories or logs', () => {
  const c = fixture('{}');
  c.SpreadsheetApp = {getActiveSpreadsheet: () => ({})};
  c.saveLineUserDirectory_ = c.saveLiffOperationLog_ = () => { throw Error('unexpected write'); };
  c.getImportantInfoLinks_ = () => ({pageUrl: 'https://example.test/info'});
  c.getLiffUserListFast_ = () => [];
  c.findUserDirectoryMatchByLineOrName_ = () => ({name:'user',userId:'test'});
  for (const cached of [true, false]) {
    c.getLiffInitDataFromDisplayMaster_ = () => cached ? {staffName:'staff',users:[]} : null;
    c.getStaffNameCached_ = () => 'staff';
    assert.equal(c.initLiffApp_('id','name').success, true);
    assert.equal(c.getImportantInfoForLiff_('id','name').importantInfo.pageUrl, 'https://example.test/info');
  }
  c.getStaffNameCached_ = () => '未登録';
  assert.equal(c.getImportantInfoForLiff_('id','name').role, 'user');
});

test('unknown important-info visitors remain denied and recorded', () => {
  const c = fixture('{}'); let writes = 0;
  c.SpreadsheetApp = {getActiveSpreadsheet: () => ({})};
  c.getLiffInitDataFromDisplayMaster_ = () => null;
  c.getStaffNameCached_ = () => '未登録';
  c.findUserDirectoryMatchByLineOrName_ = () => null;
  c.saveLineUserDirectory_ = c.saveUnregisteredLiffLogin_ = () => { writes++; };
  c.getImportantInfoLinks_ = () => {throw Error('must not return links');};
  assert.equal(c.getImportantInfoForLiff_('id','name').success, false);
  assert.equal(writes, 2);
});
