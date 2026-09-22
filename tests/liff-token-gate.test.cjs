const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const STAFF = 'U' + '1'.repeat(32);
const OTHER = 'U' + '2'.repeat(32);

function backend() {
  const c = vm.createContext({});
  for (const name of ['コード.js', 'UserLink.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas', name), 'utf8'), c);
  }
  const store = {};
  c.CacheService = { getScriptCache: () => ({ get: k => store[k] || null, put: (k, v) => { store[k] = v; } }) };
  c.Utilities = { DigestAlgorithm: { SHA_256: 'sha' }, computeDigest: (_a, t) => Array.from(t).map(ch => ch.charCodeAt(0)), base64EncodeWebSafe: b => b.join('.') };
  c.liffResponse_ = (_e, data) => data;
  let seen = null;
  c.getLiffUserList_ = id => { seen = id; return { success: true, lineUserId: id }; };
  return { c, store, seen: () => seen };
}

test('requests without a LINE-verified token are refused before any handler runs', () => {
  const b = backend();
  b.c.userLinkVerify_ = () => { const e = new Error('x'); e.userLinkSafe = true; throw e; };
  const noToken = b.c.doGet({ parameter: { action: 'getUsers', lineUserId: STAFF } });
  assert.equal(noToken.authRequired, true);
  const badToken = b.c.doGet({ parameter: { action: 'getUsers', lineUserId: STAFF, accessToken: 'forged' } });
  assert.equal(badToken.authRequired, true);
  assert.equal(b.seen(), null, 'the handler never saw the claimed ID');
});

test('a valid token for a different LINE account than the claimed lineUserId is refused', () => {
  const b = backend();
  b.c.userLinkVerify_ = () => ({ userId: OTHER, displayName: 'someone' });
  const result = b.c.doGet({ parameter: { action: 'getUsers', lineUserId: STAFF, accessToken: 'real-token-of-other' } });
  assert.equal(result.authRequired, true);
  assert.equal(b.seen(), null);
});

test('handlers receive the verified LINE ID, and a verified token is cached', () => {
  const b = backend();
  let verifications = 0;
  b.c.userLinkVerify_ = () => { verifications++; return { userId: STAFF, displayName: '佐藤' }; };
  const first = b.c.doGet({ parameter: { action: 'getUsers', accessToken: 'tok' } });
  assert.equal(first.lineUserId, STAFF, 'ID comes from LINE even when the page sent none');
  b.c.doGet({ parameter: { action: 'getUsers', lineUserId: STAFF, accessToken: 'tok' } });
  assert.equal(verifications, 1, 'second request is served from cache');
});

test('ping stays public; the LINE webhook and user POST routes keep their own checks', () => {
  const b = backend();
  b.c.userLinkVerify_ = () => { throw new Error('ping must not verify'); };
  b.c.PRICING_POLICY_VERSION = 'x';
  assert.equal(b.c.doGet({ parameter: { action: 'ping' } }).success, true);
  const src = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  const doPost = src.slice(src.indexOf('function doPost('), src.indexOf('return doGet({', src.indexOf('function doPost(')));
  assert.match(doPost, /userLinkSchedules/);
  assert.match(doPost, /userSelectFirstVisit/);
});

test('the page sends its LIFF access token on every server request and no longer uses JSONP', () => {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  assert.doesNotMatch(html, /jsonp\(/);
  const start = html.indexOf('async function requestLiffData');
  const end = html.indexOf('function logLogin', start);
  let sentBody = null;
  const c = vm.createContext({
    GAS_URL: 'https://example.com', AbortController, setTimeout, clearTimeout, URLSearchParams,
    window: {}, liff: { getAccessToken: () => 'liff-token' },
    fetch: (_url, options) => { sentBody = JSON.parse(options.body); return Promise.resolve({ ok: true, json: async () => ({}) }); }
  });
  c.window.liff = c.liff;
  vm.runInContext(html.slice(start, end), c);
  return c.requestLiffData(new URLSearchParams({ action: 'getSchedules', lineUserId: STAFF })).then(() => {
    assert.equal(sentBody.accessToken, 'liff-token');
    assert.equal(sentBody.action, 'getSchedules');
  });
});
