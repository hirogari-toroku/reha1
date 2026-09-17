const {chromium} = require('playwright');
const fs = require('node:fs'), path = require('node:path'), assert = require('node:assert/strict');
(async () => {
  const browser = await chromium.launch({headless:true,channel:'chrome'});
  try {
    for (const mode of ['success','failure','timing']) {
      const success = mode !== 'failure';
      const page = await browser.newPage({viewport:{width:390,height:844}});
      const calls = [], opened = [];
      await page.exposeFunction('recordOpen', url => opened.push(url));
      await page.route('https://hirogari-toroku.github.io/**', r => r.fulfill({contentType:'text/html',body:fs.readFileSync(path.join(__dirname,'../index.html'),'utf8')}));
      await page.route('https://static.line-scdn.net/**', r => r.fulfill({contentType:'application/javascript',body:`window.liff={init:async()=>{},isLoggedIn:()=>true,getProfile:async()=>({userId:'U${'1'.repeat(32)}',displayName:'テスト'}),isInClient:()=>true,openWindow:({url})=>window.recordOpen(url)};`}));
      await page.route('https://script.google.com/**', async r => {
        const u = new URL(r.request().url()); calls.push(u.searchParams.get('action'));
        await r.fulfill({contentType:'application/javascript',body:`${u.searchParams.get('callback')}(${JSON.stringify({success,message:success?'保存しました':'保存失敗'})})`});
      });
      await page.goto('https://hirogari-toroku.github.io/reha1/?mode=register' + (mode === 'timing' ? '&timing=1' : ''));
      await page.waitForFunction(() => !document.getElementById('registerOpenForm').disabled);
      if (mode === 'timing') await page.getByText(/ここまで合計/).waitFor();
      else if (success) await page.waitForFunction(() => document.getElementById('message').innerText.includes('保存しました'));
      await page.waitForTimeout(150);
      assert.deepEqual(calls,['logRegisterLiffLogin']);
      assert.equal(opened.length,mode === 'success'?1:0);
      if(mode === 'success') assert.match(opened[0],/^https:\/\/docs.google.com\/forms\//);
      await page.close();
    }
    console.log('Registration sends one GAS request; opens form on success only');
  } finally { await browser.close(); }
})().catch(e => {console.error(e);process.exitCode=1;});
