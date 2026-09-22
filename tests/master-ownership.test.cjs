const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
function context() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  return c;
}
test('questionnaire importer does not match LINE identities by form name', () => {
  const source = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  const importer = source.split('function importUserQuestionnaireToUserMasterCore_(ss) {')[1]
    .split('function ensureUserMasterBaseColumns_')[0];
  assert.doesNotMatch(importer, /findLineUserDirectoryByDisplayName_|lineMatch\./);
});
test('legacy user LINE writer cannot read or overwrite user master', () => {
  const c = context();
  const sheet = new Proxy({}, {get() { throw Error('legacy master accessed'); }});
  const result = c.applyLineUserDirectoryToUserRow_(sheet, {}, 2, 'family', 'msg', 'liff');
  assert.equal(result.updated, false);
  assert.equal(result.status, '利用者LINE連携で管理');
});
test('directory sync keeps staff updates but routes users to confirmed links', () => {
  const c = context();
  const row = type => ['', '', 'msg', 'liff', 'display', type, '', 2, 2];
  const directory = {getLastRow: () => 3, getRange: () => ({
    getValues: () => [row('利用者'), row('スタッフ')], setValue() {}, setValues() {}
  })};
  c.ensureLineUserDirectorySheet_ = () => directory;
  c.updateLineUserDirectoryLinksWithoutAlert_ = () => {};
  c.getStaffMasterColumnMap_ = () => ({});
  let staffWrites = 0;
  c.applyLineUserDirectoryToStaffRow_ = () => {
    staffWrites++; return {updated:true, skipped:false, status:'ok', note:''};
  };
  const result = c.applyLineUserDirectoryToMastersCore_({getSheetByName(name) {
    assert.equal(name, 'スタッフマスタ'); return {};
  }});
  assert.equal(staffWrites, 1);
  assert.equal(result.userUpdatedCount, 0);
  assert.equal(result.skippedCount, 1);
});
