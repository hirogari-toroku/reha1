const { test } = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const path = require('node:path');

// The two Playwright scripts drive index.html in a real browser with LINE and Apps
// Script stubbed out. They only run where Playwright is installed (npm install), so
// the standard suite stays runnable without it.
let playwrightAvailable = true;
try {
  require.resolve('playwright');
} catch (error) {
  playwrightAvailable = false;
}

for (const script of ['registration-browser.cjs', 'user-link-browser.cjs']) {
  test(script + ' passes in a real browser', { skip: playwrightAvailable ? false : 'playwright is not installed (run npm install)' }, () => {
    const output = execFileSync(process.execPath, [path.join(__dirname, script)], { encoding: 'utf8', timeout: 240000 });
    assert.match(output, /passed|checks passed|opens form/i, output);
  });
}
