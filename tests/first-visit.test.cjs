const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function makeSheet(rows) {
  const at = (r, col) => (rows[r - 1] || [])[col - 1];
  return {
    rows,
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
    getMaxRows: () => 1000,
    setFrozenRows() {},
    getDataRange() { return this.getRange(1, 1, rows.length, this.getLastColumn()); },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = at(row + i, col + j);
          return v === undefined ? '' : v;
        })),
      setValue: v => { while (rows.length < row) rows.push([]); rows[row - 1][col - 1] = v; },
      setValues: values => values.forEach((line, i) => {
        while (rows.length < row + i) rows.push([]);
        line.forEach((v, j) => { rows[row - 1 + i][col - 1 + j] = v; });
      })
    }),
    appendRow(values) { rows.push(values.slice()); }
  };
}

const LINK_HEADERS = ['利用者ID', '公式LINEユーザーID', '続柄', '本人家族確認', 'LIFF用LINEユーザーID', '状態', '招待ハッシュ', '有効期限', '連携日時', '確認者', '送信日時', '利用者名', '閲覧者名', 'メモ'];
const USER_LINE = 'U' + 'a'.repeat(32);
const FAMILY_LINE = 'U' + 'b'.repeat(32);
const STAFF_LINE = 'U' + 'c'.repeat(32);

function setup() {
  const sheets = {
    '利用者マスタ': makeSheet([['利用者名', '利用者ID', '状態'], ['山田太郎', 'U001', '利用中'], ['鈴木花子', 'U002', '利用中']]),
    'スタッフマスタ': makeSheet([['スタッフ名', 'LINE表示名', 'LINEユーザーID', 'LIFF用LINEユーザーID', 'スタッフID'], ['佐藤一郎', 'さとう', STAFF_LINE, 'L-sato', 'S001']]),
    '利用者LINE連携': makeSheet([
      LINK_HEADERS,
      ['U001', USER_LINE, '本人', '確認済み', 'L-user', '連携済み', '', '', '', '', '', '山田太郎', '', ''],
      ['U001', FAMILY_LINE, '長女', '確認済み', 'L-family', '連携済み', '', '', '', '', '', '山田太郎', '', ''],
      ['U001', 'U' + 'd'.repeat(32), '', '未確認', 'L-unknown', '未連携', '', '', '', '', '', '', '', '']
    ]),
    '訪問予定': makeSheet([['登録日時', 'スタッフ名', '利用者名', '訪問日', '', '', '', '', '状態', '更新日時', '更新内容']])
  };
  const ss = {
    getSheetByName: name => sheets[name] || null,
    insertSheet: name => { sheets[name] = makeSheet([]); return sheets[name]; }
  };
  const pushes = [];
  const c = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/UserLink.js'), 'utf8'), c);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/FirstVisit.js'), 'utf8'), c);
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss, flush() {} };
  c.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  c.Utilities = { formatDate: () => '20260922120000' };
  c.isAdminLiffUser_ = (_ss, id) => id === 'ADMIN';
  c.adminDeniedResponse_ = () => ({ success: false, message: 'denied' });
  c.sendPushMessage_ = (_ss, to, _name, text) => { pushes.push({ to, text }); return { success: true }; };
  return { c, ss, sheets, pushes };
}

function futureLocal(daysAhead, hour) {
  const d = new Date();
  d.setDate(d.getDate() + daysAhead);
  const pad = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + 'T' + pad(hour) + ':00';
}

test('admin creates an offer and it is pushed only to confirmed user/family accounts', () => {
  const { c, sheets, pushes } = setup();
  const result = c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', {
    userId: 'U001', staffId: 'S001', option1: futureLocal(5, 14), option2: futureLocal(3, 10), option3: ''
  });
  assert.equal(result.success, true, result.message);
  const offerRows = sheets['初回訪問候補'].rows;
  assert.equal(offerRows.length, 2);
  assert.equal(offerRows[1][9], '候補提示中');
  assert.ok(offerRows[1][6] < offerRows[1][7], 'options are stored earliest first');
  assert.deepEqual(pushes.map(p => p.to).sort(), [USER_LINE, FAMILY_LINE].sort());
  assert.match(pushes[0].text, /https:\/\/liff\.line\.me\/2010856600-yxOaV2Np/);
  assert.equal(offerRows[1][14], 'LINE送信 2/2件');
});

test('invalid requests are refused before anything is written or sent', () => {
  const { c, sheets, pushes } = setup();
  assert.equal(c.adminCreateFirstVisitCandidatesFromLiff_('someone', '', {}).success, false);
  assert.equal(c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S999', option1: futureLocal(2, 10) }).success, false);
  assert.match(c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S001', option1: '2020-01-01T10:00' }).message, /過去/);
  assert.match(c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S001' }).message, /1つ以上/);
  assert.equal(sheets['初回訪問候補'], undefined);
  assert.equal(pushes.length, 0);
});

test('a new offer replaces the open one; the user sees only the latest future options', () => {
  const { c, ss, sheets } = setup();
  c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S001', option1: futureLocal(2, 10) });
  c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S001', option1: futureLocal(4, 11), option2: futureLocal(6, 15) });
  const rows = sheets['初回訪問候補'].rows;
  assert.equal(rows[1][9], '差し替え');
  assert.equal(rows[2][9], '候補提示中');
  rows[2][0] = 'FV-latest';

  const offer = c.getFirstVisitOfferForUser_(ss, { id: 'U001', name: '山田太郎' });
  assert.equal(offer.candidateId, 'FV-latest');
  assert.equal(offer.options.length, 2);
  assert.equal(offer.staffName, '佐藤一郎');
  assert.equal(c.getFirstVisitOfferForUser_(ss, { id: 'U002', name: '鈴木花子' }), null, 'other users see nothing');
});

test('picking an option books a schedule row once and notifies staff and admin', () => {
  const { c, ss, sheets, pushes } = setup();
  c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S001', option1: futureLocal(3, 10), option2: futureLocal(4, 13) });
  const candidateId = sheets['初回訪問候補'].rows[1][0];
  pushes.length = 0;

  const user = { id: 'U001', name: '山田太郎' };
  const result = c.bookFirstVisitOption_(ss, user, 'L-family', candidateId, 2);
  assert.equal(result.success, true, result.message);
  const schedule = sheets['訪問予定'].rows;
  assert.equal(schedule.length, 2);
  assert.equal(schedule[1][1], '佐藤一郎');
  assert.equal(schedule[1][2], '山田太郎');
  assert.match(String(schedule[1][3]), /^\d{1,2}\/\d{1,2}$/);
  assert.equal(schedule[1][8], '予定');
  assert.match(schedule[1][10], /初回訪問 13:00/);
  const offer = sheets['初回訪問候補'].rows[1];
  assert.equal(offer[9], '確定');
  assert.equal(offer[13], 2);
  assert.deepEqual(pushes.map(p => p.to), [STAFF_LINE, 'Uc21fa34144f5bc50c6e5324d5e4de344']);

  assert.throws(() => c.bookFirstVisitOption_(ss, user, 'L-family', candidateId, 1), /確定済み/);
  assert.throws(() => c.bookFirstVisitOption_(ss, { id: 'U002', name: '鈴木花子' }, 'L-x', candidateId, 1), /見つかりません/);
  assert.equal(sheets['訪問予定'].rows.length, 2, 'no second booking');
});

test('user selection requires a verified, admin-confirmed LINE login', () => {
  const { c } = setup();
  c.userLinkVerify_ = () => ({ userId: 'L-unknown' });
  const denied = c.userSelectFirstVisitRequest_({ accessToken: 't', candidateId: 'x', optionIndex: 1 });
  assert.equal(denied.success, false);
  c.userLinkVerify_ = () => { const e = new Error('LINEでログインし直してください。'); e.userLinkSafe = true; throw e; };
  assert.match(c.userSelectFirstVisitRequest_({}).message, /ログイン/);
});

test('dates read back from the sheet as Date cells still parse', () => {
  const { c } = setup();
  const d = new Date(2026, 9, 1, 10, 30);
  assert.equal(c.parseFirstVisitOption_(d).getTime(), d.getTime());
  assert.equal(c.parseFirstVisitOption_('2026-02-30T10:00'), null);
  assert.equal(c.formatFirstVisitOption_(d), '10/1（木） 10:30');
});

test('frontend wires the admin form and the user pick buttons', () => {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  assert.match(html, /runAdminAction\("adminCreateFirstVisitCandidates"/);
  assert.match(html, /action: "userSelectFirstVisit", accessToken: liff\.getAccessToken\(\)/);
  assert.match(html, /if \(userFirstVisitOffer\) area\.appendChild\(buildFirstVisitOfferCard/);
  const code = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  assert.match(code, /json\.action === "userSelectFirstVisit"/);
  assert.match(code, /action === "adminCreateFirstVisitCandidates"/);
});

test('an unanswered offer can be withdrawn, and only that', () => {
  const { c, ss, sheets } = setup();
  c.adminCreateFirstVisitCandidatesFromLiff_('ADMIN', '', { userId: 'U001', staffId: 'S001', option1: futureLocal(3, 10) });
  const candidateId = sheets['初回訪問候補'].rows[1][0];

  assert.equal(c.adminCancelFirstVisitCandidatesFromLiff_('someone', '', candidateId).success, false, 'non-admins are refused');
  assert.equal(c.adminCancelFirstVisitCandidatesFromLiff_('ADMIN', '', 'FV-missing').success, false);

  assert.equal(c.adminCancelFirstVisitCandidatesFromLiff_('ADMIN', '', candidateId).success, true);
  assert.equal(sheets['初回訪問候補'].rows[1][9], '取消');
  assert.equal(c.getFirstVisitOfferForUser_(ss, { id: 'U001', name: '山田太郎' }), null, 'the user can no longer pick it');
  assert.equal(c.getOpenFirstVisitCandidatesForAdmin_(ss).length, 0, 'and it leaves the admin list');
  assert.throws(() => c.bookFirstVisitOption_(ss, { id: 'U001', name: '山田太郎' }, 'L-user', candidateId, 1), /確定済み|返事待ち|見つかりません/);

  assert.match(c.adminCancelFirstVisitCandidatesFromLiff_('ADMIN', '', candidateId).message, /返事待ちではありません/);
});

test('the admin screen offers the withdraw button and routes it', () => {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  assert.match(html, /runAdminAction\("adminCancelFirstVisitCandidates", \{ candidateId: item\.candidateId \}/);
  const code = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  assert.match(code, /action === "adminCancelFirstVisitCandidates"/);
});
