const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const c = vm.createContext({});
vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/コード.js'), 'utf8'), c);
const ymd = d => [d.getFullYear(), d.getMonth() + 1, d.getDate()].join('-');

test('M/d dates take the year closest to the base date', () => {
  assert.equal(ymd(c.parseComparisonDate_('1/5', new Date(2026, 11, 20))), '2027-1-5', 'January visit booked in December');
  assert.equal(ymd(c.parseComparisonDate_('12/28', new Date(2027, 0, 3))), '2026-12-28', 'late-December visit recorded in January');
  assert.equal(ymd(c.parseComparisonDate_('9/22', new Date(2026, 8, 22, 15))), '2026-9-22');
  assert.equal(ymd(c.parseComparisonDate_('4/1', new Date(2026, 8, 1))), '2027-4-1', 'a booking made in September for April is next April');
  assert.equal(ymd(c.parseComparisonDate_('7/10', new Date(2026, 8, 1))), '2026-7-10', 'a back-dated date within 4 months stays in the past');
  assert.equal(ymd(c.parseComparisonDate_('3月1日', new Date(2026, 1, 20))), '2026-3-1');
  assert.equal(ymd(c.parseComparisonDate_('2026/1/5', new Date(2026, 11, 20))), '2026-1-5', 'explicit years are untouched');
});
