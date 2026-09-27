const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function makeSheet(rows) {
  return {
    rows,
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
    getDataRange() { return this.getRange(1, 1, rows.length, this.getLastColumn()); },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = (rows[row - 1 + i] || [])[col - 1 + j];
          return v === undefined ? '' : v;
        }))
    })
  };
}

const LINK_HEADERS = ['利用者ID', '公式LINEユーザーID', '続柄', '本人家族確認', 'LIFF用LINEユーザーID', '状態', '招待ハッシュ', '有効期限', '連携日時', '確認者', '送信日時', '利用者名', '閲覧者名', 'メモ'];
const USER_LINE = 'U' + 'a'.repeat(32);
const FAMILY_LINE = 'U' + 'b'.repeat(32);

function setup(balances) {
  const sheets = {
    '回数券管理': makeSheet([['利用者名', '回数券残数']].concat(balances)),
    '利用者マスタ': makeSheet([['利用者名', '利用者ID', '状態'], ['山田太郎', 'U001', '利用中'], ['鈴木花子', 'U002', '利用中'], ['佐藤次郎', 'U003', '利用中']]),
    '利用者LINE連携': makeSheet([
      LINK_HEADERS,
      ['U001', USER_LINE, '本人', '確認済み', 'L1', '連携済み', '', '', '', '', '', '山田太郎', '', ''],
      ['U001', FAMILY_LINE, '長女', '確認済み', 'L2', '連携済み', '', '', '', '', '', '山田太郎', '', ''],
      ['U002', 'U' + 'c'.repeat(32), '', '未確認', 'L3', '未連携', '', '', '', '', '', '鈴木花子', '', '']
    ])
  };
  const ss = { getSheetByName: name => sheets[name] || null };
  const c = vm.createContext({ console });
  for (const name of ['コード.js', 'UserLink.js', 'FirstVisit.js', 'CouponNotice.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas', name), 'utf8'), c);
  }
  const props = {};
  c.PropertiesService = { getScriptProperties: () => ({
    getProperty: k => (k in props ? props[k] : null),
    setProperty: (k, v) => { props[k] = v; }
  }) };
  c.Utilities = { formatDate: () => '9月27日' };
  const pushes = [];
  c.sendPushMessage_ = (_ss, to, _name, text) => { pushes.push({ to, text }); return { success: true }; };
  return { c, ss, pushes, props };
}

test('the visit that uses the last coupon tells linked user and family accounts, and the staff reply', () => {
  const { c, ss, pushes } = setup([['山田太郎', 1]]);
  const note = c.couponNoticeOnVisitStart_(ss, '山田太郎', '9/25');
  assert.deepEqual(pushes.map(p => p.to), [USER_LINE, FAMILY_LINE]);
  assert.match(pushes[0].text, /^山田太郎 様\n/);
  assert.match(pushes[0].text, /本日（9月25日）の訪問で、お手元の回数券をすべてご利用いただくことになります。/);
  assert.match(pushes[0].text, /お振込み先はいつもの口座です。/);
  assert.match(note, /山田太郎様の回数券は今回で使い切りになります。/);
  assert.match(note, /ご案内をお願いいたします。$/);
});

test('a visit with no coupons left says it is unpaid and how many visits are owed', () => {
  const { c, ss, pushes } = setup([['山田太郎', -1]]);
  const note = c.couponNoticeOnVisitStart_(ss, '山田太郎', '9/25');
  assert.equal(pushes.length, 2);
  assert.match(pushes[0].text, /本日（9月25日）の訪問は未精算となります（本日分を含めて未精算 2回分）/);
  assert.match(note, /残りがなく、今回は未精算になります。/);

  const zero = setup([['山田太郎', 0]]);
  zero.c.couponNoticeOnVisitStart_(zero.ss, '山田太郎', '9/25');
  assert.match(zero.pushes[0].text, /未精算 1回分/);
});

test('nothing is sent while coupons remain, and no admin message is ever sent', () => {
  const { c, ss, pushes } = setup([['山田太郎', 2]]);
  assert.equal(c.couponNoticeOnVisitStart_(ss, '山田太郎', '9/25'), '');
  assert.equal(pushes.length, 0);
});

test('an unlinked user gets no LINE but the staff still sees the note', () => {
  const { c, ss, pushes } = setup([['鈴木花子', 1], ['佐藤次郎', 0]]);
  assert.match(c.couponNoticeOnVisitStart_(ss, '鈴木花子', '9/25'), /使い切り/);
  assert.match(c.couponNoticeOnVisitStart_(ss, '佐藤次郎 様', '9/25'), /未精算/);
  assert.equal(pushes.length, 0);
});

test('recording the start again on the same day does not send the notice twice', () => {
  const { c, ss, pushes } = setup([['山田太郎', 1]]);
  c.couponNoticeOnVisitStart_(ss, '山田太郎', '9/25');
  const again = c.couponNoticeOnVisitStart_(ss, '山田太郎', '9/25');
  assert.equal(pushes.length, 2);
  assert.match(again, /使い切り/, 'the staff reply still carries the note');
  c.couponNoticeOnVisitStart_(ss, '山田太郎', '10/2');
  assert.equal(pushes.length, 4, 'the next visit day is a new notice');
});

test('both ways of recording a start carry the note, and the end does not', () => {
  const code = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  assert.match(code, /firstRow\[1\] === "開始" \? couponNoticeOnVisitStart_\(ss, firstRow\[3\], firstRow\[4\]\) : ""/);
  assert.match(code, /type === "開始" \? couponNoticeOnVisitStart_\(ss, resolvedUserName, targetDate\) : ""/);
  assert.match(code, /"利用者：" \+ firstRow\[3\] \+ couponNote/);
});

test('the test user is never notified', () => {
  const { c, ss, pushes } = setup([['テスト利用者', 0]]);
  assert.equal(c.couponNoticeOnVisitStart_(ss, 'テスト利用者', '9/25'), '');
  assert.equal(pushes.length, 0);
});

test('the unpaid count grows by one with each visit while no payment arrives', () => {
  [[0, 1], [-1, 2], [-2, 3], [-3, 4]].forEach(([balance, owed]) => {
    const { c, ss, pushes } = setup([['山田太郎', balance]]);
    c.couponNoticeOnVisitStart_(ss, '山田太郎', '10/2');
    assert.match(pushes[0].text, new RegExp('本日分を含めて未精算 ' + owed + '回分'));
  });
});

test('the transfer account comes from a script property, not the code', () => {
  const plain = setup([['山田太郎', 1]]);
  plain.c.couponNoticeOnVisitStart_(plain.ss, '山田太郎', '9/25');
  assert.match(plain.pushes[0].text, /お振込み先はいつもの口座です。/);

  const withAccount = setup([['山田太郎', 1]]);
  withAccount.props.COUPON_TRANSFER_ACCOUNT = '〇〇銀行 〇〇支店 普通 0000000 口座名義';
  withAccount.c.couponNoticeOnVisitStart_(withAccount.ss, '山田太郎', '9/25');
  assert.match(withAccount.pushes[0].text, /【お振込み先】\n〇〇銀行 〇〇支店 普通 0000000 口座名義\n/);
  assert.doesNotMatch(withAccount.pushes[0].text, /いつもの口座/);
});
