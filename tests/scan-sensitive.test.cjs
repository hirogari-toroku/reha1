const { test } = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const scanner = path.join(__dirname, '../tools/scan-sensitive.cjs');

// Runs the scanner over a throwaway git repo containing one file.
function scan(content, name = 'sample.md') {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'scan-'));
  try {
    execFileSync('git', ['init', '-q'], { cwd: dir });
    fs.mkdirSync(path.join(dir, 'tools'), { recursive: true });
    fs.copyFileSync(scanner, path.join(dir, 'tools/scan-sensitive.cjs'));
    fs.writeFileSync(path.join(dir, name), content);
    execFileSync('git', ['add', '-A'], { cwd: dir });
    try {
      const out = execFileSync(process.execPath, ['tools/scan-sensitive.cjs', '--all'], { cwd: dir, encoding: 'utf8' });
      return { code: 0, out: out };
    } catch (error) {
      return { code: error.status, out: (error.stdout || '') + (error.stderr || '') };
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('the repository as it stands is clean', () => {
  const out = execFileSync(process.execPath, [scanner, '--all'], { cwd: path.join(__dirname, '..'), encoding: 'utf8' });
  assert.match(out, /見つかりませんでした/);
});

test('secrets are caught', () => {
  assert.equal(scan('const token = "GOCSPX-abcdefghij1234567890";').code, 1);
  assert.equal(scan('ghp_' + 'a'.repeat(36)).code, 1);
  assert.equal(scan('-----BEGIN RSA PRIVATE KEY-----').code, 1);
  assert.equal(scan('channel_secret: "0123456789abcdef"').code, 1);
  assert.equal(scan('口座番号: 1234567').code, 1);
  assert.equal(scan('U' + '0'.repeat(32)).code, 1, 'an unknown LINE user ID is flagged');
});

test('real-looking personal names are caught, including ones close to the test samples', () => {
  const found = scan('本日、佐藤太郎様の訪問を実施しました。');
  assert.equal(found.code, 1);
  assert.match(found.out, /氏名らしい記述/);
  assert.equal(scan('架空田さんの給与を更新').code, 1);
  assert.equal(scan('佐藤花子様へ連絡').code, 1, 'a name ending like a sample name is still flagged');
});

test('generic wording, the existing admin ID and the test sample names pass', () => {
  assert.equal(scan('利用者様とご家族様へ、スタッフさんが説明します。皆様よろしくお願いします。').code, 0);
  assert.equal(scan('const ADMIN_LINE_USER_ID = "Uc21fa34144f5bc50c6e5324d5e4de344";').code, 0);
  assert.equal(scan("const text = '山田太郎様 終了';").code, 0);
});
