const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  const store = {};
  c.CacheService = { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; } }) };
  return { c, store };
}

test('cached reads are served from cache until the writer bumps the version', () => {
  const { c } = context();
  let loads = 0;
  let value = { balance: 3 };
  const read = () => JSON.parse(JSON.stringify(c.cachedRead_('couponDisplay', () => { loads++; return value; })));
  assert.deepEqual(read(), { balance: 3 });
  assert.deepEqual(read(), { balance: 3 });
  assert.equal(loads, 1);
  value = { balance: 2 };
  c.bumpReadCache_('couponDisplay');
  assert.deepEqual(read(), { balance: 2 });
  assert.equal(loads, 2);
});

test('an evicted version starts a new one instead of reusing old cached data', () => {
  const { c, store } = context();
  let value = 'old';
  c.cachedRead_('userResources', () => value);
  Object.keys(store).filter(k => k.startsWith('readCacheVer:')).forEach(k => delete store[k]);
  value = 'new';
  assert.equal(c.cachedRead_('userResources', () => value), 'new');
});

test('coupon writes bump the coupon cache', () => {
  const src = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  const body = src.slice(src.indexOf('function updateCouponManagementLocked_('), src.indexOf('function getCouponDisplayMap_('));
  assert.equal((body.match(/bumpReadCache_\("couponDisplay"\)/g) || []).length, 2);
});
