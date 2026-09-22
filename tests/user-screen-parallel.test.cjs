const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');

test('the booking request can start before the LINE profile arrives', () => {
  const start = html.indexOf('let scheduleLoadPromise = null;');
  const end = html.indexOf('function logRegisterLiffLogin', start);
  let fetched = 0;
  const c = vm.createContext({
    GAS_URL: 'https://example.com', AbortController, setTimeout, clearTimeout,
    liff: { getAccessToken: () => 'tok' }, lineUserId: '',
    document: { getElementById: () => ({ innerHTML: '', innerText: '' }) },
    setMessage: () => {}, reportLiffStage: () => {}, renderUserScheduleList: () => {},
    setResourceData: () => {}, noteSlowRequest: () => {}, liffTraceId: 'trace-id-1234',
    fetch: () => { fetched++; return Promise.resolve({ ok: true, json: async () => ({ success: true, schedules: [] }) }); }
  });
  vm.runInContext(html.slice(start, end), c);

  c.loadUserSchedules(true);
  assert.equal(fetched, 0, 'without the flag a missing ID still skips the request');
  const request = c.loadUserSchedules(true, true);
  assert.equal(fetched, 1, 'with the flag it runs while the profile is still loading');
  return request;
});

test('init starts that request in parallel and reuses it after the profile resolves', () => {
  assert.match(html, /const profilePromise = liff\.getProfile\(\);\s*\n\s*if \(appMode === "user"\) userLoadStartedEarly = loadUserSchedules\(true, true\);/);
  assert.match(html, /if \(appMode === "user"\) return userLoadStartedEarly \|\| loadUserSchedules\(true\);/);
});
