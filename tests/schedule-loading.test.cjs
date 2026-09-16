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
  assert.equal(items[0].basicInfoUrl, undefined);
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

test('user filter skips other users before expensive visit processing', () => {
  const c = backend(), parsed = [];
  c.parseComparisonDate_ = value => { parsed.push(value); return null; };
  const sheet = { getLastRow: () => 4, getLastColumn: () => 8,
    getRange: () => ({ getValues: () => [[null, '終了', 'A', 'U', 'a'], [null, '終了', 'B', 'V', 'b'], [null, '終了', 'B', 'U', 'c']] }) };
  c.buildVisitStatusIndex_(sheet, '', 'U');
  assert.deepEqual(parsed, ['a', 'c']);
});

test('staff resources expose one folder and old user JSON retains the chart', () => {
  const c = backend();
  c.getStaffMasterColumnMap_ = () => ({ name: 0, payrollFolderId: 1 });
  const sheet = { getLastRow: () => 2, getLastColumn: () => 2,
    getRange: () => ({ getValues: () => [['S', 'folder-id']] }) };
  const resource = c.getStaffResourceMap_(sheet).S;
  assert.equal(resource.staffFolderUrl, 'https://drive.google.com/drive/folders/folder-id');
  assert.equal(resource.payslipFolderUrl, undefined);
  const users = c.parseLiffUserLinksJson_(JSON.stringify([{ name: 'U', chartUrl: 'https://example.com/chart', basicInfoUrl: '' }]), '');
  assert.equal(users[0].chartUrl, 'https://example.com/chart');
  assert.equal(users[0].basicInfoUrl, undefined);
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  assert.ok(html.includes('makeResourceLink("カルテ", resources.chartUrl, true)'));
  assert.ok(!html.includes('appendUserResourceLinks'));
});

test('important-info panel is hidden for staff but retained on dedicated and user pages', () => {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const start = html.indexOf('function renderImportantInfoLinks()');
  const end = html.indexOf('function getUserResourceForItem', start);
  for (const mode of ['staff', 'user', 'important', 'admin']) {
    let hidden = false;
    const children = [];
    const area = { innerHTML: '', appendChild: item => children.push(item) };
    const panel = { classList: { add: () => { hidden = true; }, remove: () => { hidden = false; } } };
    const c = vm.createContext({ appMode: mode,
      importantInfo: { consentFormUrl: 'https://example.com/consent' },
      document: { getElementById: id => id === 'importantInfoPanel' ? panel : area },
      isValidResourceUrl: url => Boolean(url), makeResourceLink: label => label });
    vm.runInContext(html.slice(start, end), c);
    c.renderImportantInfoLinks();
    assert.equal(hidden, mode === 'staff' || mode === 'admin');
    assert.equal(children.length, hidden ? 0 : 1);
  }
});

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
