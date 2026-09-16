const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function fixture(extraUsers = [], extraPairs = []) {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  c.Utilities = { formatDate: (_date, _tz, format) => format === 'M/d' ? '9/16' : '16:00' };
  c.findLastStartedUser_ = () => { throw Error('must not infer previous user'); };
  c.resolveUserName_ = () => { throw Error('must not resolve surnames'); };
  const ss = { getSheetByName: name => ({ getDataRange: () => ({ getValues: () =>
    name === 'スタッフ利用者マスタ'
      ? [[], ['担当A', '山田太郎', '山田'], ['担当A', 'テスト利用者'], ...extraPairs]
      : [[], ['山田太郎'], ...extraUsers]
  }) }) };
  return text => c.parseVisitResult_(ss, text, new Date(), '担当A', 'id');
}

test('only assigned full-name commands register, preserving canonical name', () => {
  const parse = fixture();
  for (const text of ['山田太郎 開始', '山田 太郎　開始', '山田太郎 終了']) {
    const rows = parse(text);
    assert.equal(rows.length, 1);
    assert.equal(rows[0][3], '山田太郎');
  }
  assert.equal(parse('テスト利用者 開始').length, 1);
});

test('surname, absent names, honorifics, aliases and conversational messages never register', () => {
  const parse = fixture();
  for (const text of ['山田 開始', '開始', '終了', '山田様 開始', '山田太郎様 開始', '山田太郎さん 終了',
    '山田太郎開始', '山田太郎 開始しました', '山田太郎は来週開始予定です',
    '山田太郎について相談 開始', '山田太郎 開始？', '山田太郎\n開始',
    '山田太郎 開始\n相談です', '「山田太郎 開始」', '佐藤次郎 終了']) {
    assert.equal(parse(text).length, 0, text);
  }
});

test('ambiguous duplicate master names or assignments are rejected', () => {
  assert.equal(fixture([['山田太郎']])('山田太郎 開始').length, 0);
  assert.equal(fixture([], [['担当A', '山田太郎']])('山田太郎 開始').length, 0);
});
