const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function fixture() {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const start = html.indexOf('function runAdminAction');
  const end = html.indexOf('function renderScheduleList', start);
  let resolveJsonp, rejectJsonp;
  const scrolled = { count: 0 };
  const messageEl = { innerText: '', scrollIntoView: () => { scrolled.count++; } };
  const c = vm.createContext({
    URLSearchParams, lineUserId: 'admin-id', lineDisplayName: '管理者', GAS_URL: 'https://example.com',
    document: { getElementById: id => id === 'message' ? messageEl : null },
    setMessage: text => { messageEl.innerText = text; },
    requestLiffData: () => new Promise((yes, no) => { resolveJsonp = yes; rejectJsonp = no; }),
    loadAdminDashboard: () => Promise.resolve()
  });
  vm.runInContext(html.slice(start, end), c);
  return { c, messageEl, scrolled,
    resolveJsonp: value => resolveJsonp(value),
    rejectJsonp: error => rejectJsonp(error) };
}

const flush = () => new Promise(resolve => setTimeout(resolve, 10));

test('a data-maintenance button shows immediate feedback and restores after success', async () => {
  const f = fixture();
  const button = { disabled: false, textContent: '利用者アンケート取込' };

  f.c.runAdminAction('adminImportUserQuestionnaire', null, button);

  assert.equal(button.disabled, true, 'クリック直後にボタンを無効化するべき');
  assert.equal(button.textContent, '処理中...', 'クリック直後にラベルを変えるべき');
  assert.equal(f.messageEl.innerText, '処理中...');
  assert.ok(f.scrolled.count >= 1, 'メッセージが見える位置までスクロールされるべき');

  f.resolveJsonp({ success: true, message: '取り込みました' });
  await flush();

  assert.equal(button.disabled, false, '完了後はボタンを元に戻すべき');
  assert.equal(button.textContent, '利用者アンケート取込', 'ラベルを元の文言に戻すべき');
  assert.equal(f.messageEl.innerText, '取り込みました');
});

test('a failed action still restores the button and shows an error message', async () => {
  const f = fixture();
  const button = { disabled: false, textContent: '回数券更新' };

  f.c.runAdminAction('adminRunCouponAll', null, button);
  f.rejectJsonp(new Error('timeout'));
  await flush();

  assert.equal(button.disabled, false);
  assert.equal(button.textContent, '回数券更新');
  assert.match(f.messageEl.innerText, /処理に失敗しました/);
});

test('calling without a button element (legacy callers) does not throw', async () => {
  const f = fixture();
  assert.doesNotThrow(() => f.c.runAdminAction('adminSaveRelationship', { rowNumber: '2', staffId: 'S1', userId: 'U1' }));
  f.resolveJsonp({ success: true, message: '保存しました' });
  await flush();
  assert.equal(f.messageEl.innerText, '保存しました');
});
