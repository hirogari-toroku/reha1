const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function fixture() {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const start = html.indexOf('async function requestLiffData');
  const end = html.indexOf('function renderAdminDashboard', start);
  let resolve, reject, jsonpCalls = 0, fetchCalls = 0;
  const summary = { innerText: '' };
  const c = vm.createContext({
    URLSearchParams, lineUserId: 'admin-id', lineDisplayName: '管理者', GAS_URL: 'https://example.com',
    document: { getElementById: id => id === 'adminSummary' ? summary : { innerText: '' } },
    setMessage: () => {}, AbortController, clearTimeout, setTimeout,
    jsonp: () => { jsonpCalls++; return new Promise(() => {}); },
    fetch: () => { fetchCalls++; return new Promise((yes, no) => { resolve = yes; reject = no; }); }
  });
  vm.runInContext(html.slice(start, end), c);
  return { c, jsonpCalls: () => jsonpCalls, fetchCalls: () => fetchCalls, summary,
    resolve: value => resolve({ ok: true, json: async () => value }),
    reject: error => reject(error) };
}

test('admin dashboard loads over POST (fetch), not the legacy jsonp GET path', async () => {
  const f = fixture();
  f.c.renderAdminDashboard = () => {};
  const request = f.c.loadAdminDashboard(true);
  assert.equal(f.fetchCalls(), 1);
  assert.equal(f.jsonpCalls(), 0);
  f.resolve({ success: true, message: '' });
  await request;
});

test('admin dashboard failure surfaces a message without throwing', async () => {
  const f = fixture();
  const request = f.c.loadAdminDashboard(false);
  f.reject(new Error('timeout'));
  await request;
  assert.match(f.summary.innerText, /取得に失敗/);
});
