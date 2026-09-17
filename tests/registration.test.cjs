const {test} = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs'), vm = require('node:vm'), path = require('node:path');
const id = 'U' + '1'.repeat(32);
function fixture(rows = []) {
  const c = vm.createContext({});
  for (const file of ['コード.js','Registration.js']) vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',file),'utf8'), c);
  let released = 0, reads = 0;
  const sheet = {getLastRow: () => rows.length + 1,
    getRange: (r) => ({getValues() {reads++; return rows.map(x => x.slice());}, setValues(values) {rows[r-2] = values[0];}}),
    appendRow: r => rows.push(r)};
  c.ensureLineUserDirectorySheet_ = () => sheet;
  c.LockService = {getScriptLock: () => ({tryLock: () => true, releaseLock() {released++;}})};
  c.SpreadsheetApp = {flush() {}, getActiveSpreadsheet: () => ({getSheetByName() {throw Error('unexpected master read');}})};
  return {c,rows,get reads() {return reads;},get released() {return released;}};
}
test('registration captures only directory data, without granting access or master scans', () => {
  const f = fixture();
  assert.equal(f.c.logRegisterLiffLogin_(id,'表示名').success,true);
  assert.equal(f.rows[0][3],id); assert.equal(f.rows[0][4],'表示名');
  assert.equal(f.rows[0][9],'未確認'); assert.equal(f.released,1);
  f.c.logRegisterLiffLogin_(id,'新表示名');
  assert.equal(f.rows.length,1); assert.equal(f.reads,1);
});
test('registration preserves official ID, relation message, notes and existing hints', () => {
  const row = ['first','last','official',id,'old','type','user',2,3,'state','source','利用者 母','memo'];
  const f = fixture([row]); f.c.logRegisterLiffLogin_(id,'new');
  for (const i of [0,2,5,6,7,8,9,11,12]) assert.equal(f.rows[0][i],row[i]);
});
test('same display name does not merge accounts; duplicate ID and invalid inputs fail safely', () => {
  const row = Array(13).fill(''); row[4] = 'same';
  const f = fixture([row]); f.c.logRegisterLiffLogin_(id,'same'); assert.equal(f.rows.length,2);
  f.rows.push(f.rows[1].slice());
  assert.equal(f.c.logRegisterLiffLogin_(id,'same').success,false);
  assert.equal(f.c.logRegisterLiffLogin_('bad','same').success,false);
  assert.equal(f.rows.length,3);
});
