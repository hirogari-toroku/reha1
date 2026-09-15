const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function backend() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  return c;
}

test('resource reads do not run master maintenance; normal callers still do', () => {
  const c = backend();
  let maintenance = 0;
  c.ensureUserMasterBaseColumns_ = () => { maintenance++; };
  const headers = ['利用者名', '基本情報URL'];
  const sheet = {
    getLastRow: () => 2, getLastColumn: () => 2,
    getRange: row => ({ getValues: () => row === 1 ? [headers] : [['U', 'https://example.com/info']] })
  };
  const ss = { getSheetByName: () => sheet };
  const readOnly = c.getUserResourceMap_(ss, true);
  assert.equal(maintenance, 0);
  const normal = c.getUserResourceMap_(ss);
  assert.equal(maintenance, 1);
  assert.deepEqual(readOnly, normal);
});

test('schedule enrichment preserves resources and coupon without maintenance', () => {
  const c = backend();
  c.getUserResourceMap_ = (ss, readOnly) => {
    assert.equal(readOnly, true);
    return { U: { chartUrl: 'chart', basicInfoUrl: 'info' } };
  };
  c.getCouponDisplayMap_ = () => ({ U: { balance: 3 } });
  const items = [{ userName: 'U' }];
  c.enrichLiffScheduleItemsWithUserResources_({}, items);
  assert.equal(items[0].chartUrl, 'chart');
  assert.equal(items[0].basicInfoUrl, 'info');
  assert.equal(items[0].coupon.balance, 3);
});

test('staff filter skips other staff before date parsing, default keeps all', () => {
  const c = backend();
  const parsed = [];
  c.parseComparisonDate_ = value => { parsed.push(value); return null; };
  const sheet = {
    getLastRow: () => 3, getLastColumn: () => 8,
    getRange: () => ({ getValues: () => [[null, '開始', 'A', 'U', 'a'], [null, '開始', 'B', 'V', 'b']] })
  };
  c.buildVisitStatusIndex_(sheet, 'A');
  assert.deepEqual(parsed, ['a']);
  parsed.length = 0;
  c.buildVisitStatusIndex_(sheet);
  assert.deepEqual(parsed, ['a', 'b']);
});

function frontend() {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const start = html.indexOf('let scheduleLoadPromise = null;');
  const end = html.indexOf('function loadUserSchedules', start);
  let resolve, reject;
  let calls = 0, message = '';
  const c = vm.createContext({
    URLSearchParams, lineUserId: 'staff', GAS_URL: 'https://example.com',
    scheduleItems: [{ id: 'old' }],
    setMessage: value => { message = value; },
    setResourceData: () => {}, renderScheduleList: () => {},
    jsonp: (url, timeout) => {
      assert.equal(timeout, 45000);
      calls++;
      return new Promise((yes, no) => { resolve = yes; reject = no; });
    }
  });
  vm.runInContext(html.slice(start, end), c);
  return { c, calls: () => calls, message: () => message,
    resolve: value => resolve(value), reject: error => reject(error) };
}

test('concurrent schedule refreshes share one request and can refresh after completion', async () => {
  const f = frontend();
  const first = f.c.loadSchedules();
  assert.equal(f.c.loadSchedules(), first);
  assert.equal(f.calls(), 1);
  f.resolve({ success: true, schedules: [{ id: 'new' }] });
  await first;
  assert.equal(f.c.scheduleItems[0].id, 'new');
  const next = f.c.loadSchedules();
  assert.equal(f.calls(), 2);
  f.resolve({ success: true, schedules: [] });
  await next;
});

test('silent initial failures are visible and keep existing schedules with stale warning', async () => {
  const f = frontend();
  const request = f.c.loadSchedules(true);
  f.reject(new Error('timeout'));
  await request;
  assert.equal(f.c.scheduleItems[0].id, 'old');
  assert.match(f.message(), /timeout/);
  assert.match(f.message(), /更新前/);
});
