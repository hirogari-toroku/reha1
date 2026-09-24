const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');

function screen(mode, stored) {
  const start = html.indexOf('const LAST_VIEW_STORAGE_KEY');
  const end = html.indexOf('function noteSlowRequest', start);
  const store = stored === undefined ? {} : { hirogariLastView: stored };
  const rendered = [];
  const c = vm.createContext({
    appMode: mode, lineUserId: '', scheduleItems: [], userCoupon: null,
    shownLastViewAccount: null, lastShownView: null, messages: [],
    localStorage: {
      getItem: k => (k in store ? store[k] : null),
      setItem: (k, v) => { store[k] = v; },
      removeItem: k => { delete store[k]; }
    },
    setMessage: text => { c.messages.push(text); },
    renderScheduleList: () => rendered.push('staff'),
    renderUserScheduleList: () => rendered.push('user')
  });
  vm.runInContext(html.slice(start, end), c);
  return { c, store, rendered };
}

const savedUserView = JSON.stringify({
  view: 'user', lineUserId: 'U-owner', savedAt: Date.now(),
  payload: { schedules: [{ visitDate: '9/25', status: '予約あり' }], coupon: { balance: 2 } }
});

test('a saved list is shown at once with a notice, before any request finishes', () => {
  const s = screen('user', savedUserView);
  const account = s.c.showLastViewWhileLoading();
  assert.equal(account, 'U-owner');
  assert.equal(s.c.scheduleItems.length, 1);
  assert.equal(s.c.userCoupon.balance, 2);
  assert.deepEqual(s.rendered, ['user']);
  assert.match(s.c.messages[0], /前回の内容/);
});

test('nothing is shown when no copy is saved or the copy is unreadable', () => {
  assert.equal(screen('user').c.showLastViewWhileLoading(), null);
  assert.equal(screen('user', '{broken').c.showLastViewWhileLoading(), null);
  assert.equal(screen('user', JSON.stringify({ view: 'x', payload: {} })).c.showLastViewWhileLoading(), null);
});

test('a different LINE account clears the shown copy and the stored one', () => {
  const s = screen('user', savedUserView);
  s.c.shownLastViewAccount = s.c.showLastViewWhileLoading();
  s.c.lineUserId = 'U-someone-else';
  s.c.discardLastViewIfAccountChanged();
  assert.equal(s.c.scheduleItems.length, 0);
  assert.equal(s.c.userCoupon, null);
  assert.equal('hirogariLastView' in s.store, false);
  assert.match(s.c.messages[s.c.messages.length - 1], /別のLINEアカウント/);
});

test('the same account keeps the shown copy, and saving needs a known account', () => {
  const s = screen('user', savedUserView);
  s.c.shownLastViewAccount = s.c.showLastViewWhileLoading();
  s.c.lineUserId = 'U-owner';
  s.c.discardLastViewIfAccountChanged();
  assert.equal(s.c.scheduleItems.length, 1);

  const fresh = screen('staff');
  fresh.c.saveLastView({ schedules: [1] });
  assert.equal('hirogariLastView' in fresh.store, false, 'no account yet: nothing stored');
  fresh.c.lineUserId = 'U-staff';
  fresh.c.saveLastView({ schedules: [1] });
  assert.equal(JSON.parse(fresh.store.hirogariLastView).view, 'staff');
});

test('the page preconnects to LINE and Apps Script', () => {
  ['static.line-scdn.net', 'api.line.me', 'script.google.com', 'script.googleusercontent.com'].forEach(host => {
    assert.match(html, new RegExp('<link rel="preconnect" href="https://' + host.replace(/\./g, '\\.') + '"'));
  });
});
