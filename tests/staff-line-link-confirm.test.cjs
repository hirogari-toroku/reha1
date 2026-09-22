const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function context() {
  const c = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
  return c;
}

function makeSheet(rows) {
  const at = (r, col) => (rows[r - 1] || [])[col - 1];
  return {
    rows,
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
    getDataRange() { return this.getRange(1, 1, rows.length, this.getLastColumn()); },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => {
          const v = at(row + i, col + j);
          return v === undefined ? '' : v;
        })),
      setValue: v => { rows[row - 1][col - 1] = v; },
      setValues: values => values.forEach((line, i) => line.forEach((v, j) => { rows[row - 1 + i][col - 1 + j] = v; }))
    })
  };
}

const header = ['スタッフ名', 'LINE表示名', 'LINEユーザーID', 'LIFF用LINEユーザーID'];

test('directory sync that fills an empty staff LINE ID marks the row 未確認 and blocks lookup', () => {
  const c = context();
  const staff = makeSheet([header.slice(), ['佐藤花子', '', '', ''], ['田中一郎', '田中', 'U-tanaka', 'L-tanaka']]);
  const ss = { getSheetByName: () => staff };
  const cols = c.getStaffMasterColumnMap_(staff);

  const result = c.applyLineUserDirectoryToStaffRow_(staff, cols, 2, '佐藤花子', 'U-imposter', 'L-imposter');
  assert.equal(result.updated, true);
  assert.equal(staff.rows[0][4], 'LINE紐づけ確認', 'confirmation column is created on demand');
  assert.equal(staff.rows[1][4], '未確認');
  assert.match(result.note, /承認待ち/);

  assert.equal(c.getStaffName_(ss, 'U-imposter'), '未登録');
  assert.equal(c.getStaffName_(ss, 'L-imposter'), '未登録');
  assert.equal(c.getStaffName_(ss, 'U-tanaka'), '田中一郎', 'existing links (blank confirmation) keep working');
});

test('an admin approval makes the link usable; a rejection clears the auto-filled LINE columns', () => {
  const c = context();
  const rows = [header.concat(['LINE紐づけ確認']), ['佐藤花子', '佐藤', 'U-sato', 'L-sato', '未確認'], ['鈴木次郎', 'すずき', 'U-x', 'L-x', '未確認']];
  const staff = makeSheet(rows);
  const ss = { getSheetByName: name => (name === 'スタッフマスタ' ? staff : null) };
  c.SpreadsheetApp = { getActiveSpreadsheet: () => ss };
  c.isAdminLiffUser_ = (_ss, id) => id === 'ADMIN';
  c.adminDeniedResponse_ = () => ({ success: false, message: 'denied' });
  let displayRefreshes = 0;
  c.updateLiffDisplayMaster = () => { displayRefreshes++; };

  assert.equal(c.getPendingStaffLineLinks_(ss).length, 2);
  assert.equal(c.adminResolveStaffLineLinkFromLiff_('someone', '', 2, '佐藤花子', true).success, false, 'non-admins are refused');
  assert.equal(c.adminResolveStaffLineLinkFromLiff_('ADMIN', '', 2, '別人', true).success, false, 'row/name mismatch is refused');

  assert.equal(c.adminResolveStaffLineLinkFromLiff_('ADMIN', '', 2, '佐藤花子', true).success, true);
  assert.equal(rows[1][4], '確認済');
  assert.equal(c.getStaffName_(ss, 'U-sato'), '佐藤花子');

  assert.equal(c.adminResolveStaffLineLinkFromLiff_('ADMIN', '', 3, '鈴木次郎', false).success, true);
  assert.deepEqual(rows[2], ['鈴木次郎', '', '', '', '']);
  assert.equal(c.getStaffName_(ss, 'U-x'), '未登録');
  assert.equal(displayRefreshes, 2);
  assert.equal(c.getPendingStaffLineLinks_(ss).length, 0);
});

test('pending staff get no LIFF display-master entry', () => {
  const src = fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8');
  assert.match(src, /const liffLineUserId = isStaffLineLinkPending_\(staffValues\[i\], cols\)\s*\?\s*""/);
});

test('updates that do not touch an ID column leave the confirmation alone', () => {
  const c = context();
  const staff = makeSheet([header.slice(), ['佐藤花子', '', 'U-sato', 'L-sato']]);
  const cols = c.getStaffMasterColumnMap_(staff);
  const updates = [{ rowNumber: 2, column: 2, value: 'さとう' }];
  assert.equal(c.addStaffLineLinkPendingUpdate_(staff, cols, updates, 2), false);
  assert.equal(updates.length, 1);
});

test('admin screen lists pending links with approve and reject buttons', () => {
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  assert.match(html, /id="adminPendingStaffLinks"/);
  assert.match(html, /runAdminAction\("adminConfirmStaffLineLink"/);
  assert.match(html, /runAdminAction\("adminRejectStaffLineLink"/);
});
