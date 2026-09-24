const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const src = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');

function context(store) {
  const c = vm.createContext({});
  vm.runInContext(src, c);
  c.CacheService = { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; } }) };
  return c;
}

test('a sender counts as recently recorded only after the first message', () => {
  const store = {};
  const c = context(store);
  assert.equal(c.wasDirectoryRecordedRecently_('U1'), false, 'first message records the row');
  assert.equal(c.wasDirectoryRecordedRecently_('U1'), true, 'later messages within the hour skip it');
  assert.equal(c.wasDirectoryRecordedRecently_('U2'), false, 'another sender is independent');
});

test('without a usable cache the row is always recorded, and the staff name is still resolved', () => {
  const c = vm.createContext({});
  vm.runInContext(src, c);
  assert.equal(c.wasDirectoryRecordedRecently_('U1'), false);
  let reads = 0;
  c.getStaffName_ = () => { reads++; return '佐藤'; };
  assert.equal(c.getStaffNameCached_({}, 'U1'), '佐藤');
  assert.equal(c.getStaffNameCached_({}, 'U1'), '佐藤');
  assert.equal(reads, 2, 'falls back to reading the master when the cache is unavailable');
});

test('the LINE webhook uses the cached staff lookup and skips the directory write for known staff', () => {
  const webhook = src.slice(src.indexOf('function doPost('), src.indexOf('function doGet('));
  assert.match(webhook, /const staffName = getStaffNameCached_\(ss, userId\);/);
  assert.doesNotMatch(webhook, /const staffName = getStaffName_\(ss, userId\);/);
  assert.match(webhook, /if \(!isRegisteredStaff \|\| !wasDirectoryRecordedRecently_\(userId\)\) \{\s*\n\s*saveLineUserDirectory_/);
});
