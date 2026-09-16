const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');

(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'chrome' });
  try {
    for (const viewport of [{width:390,height:844},{width:1440,height:900}]) {
      const page = await browser.newPage({viewport});
      const errors=[];page.on('pageerror',e=>errors.push(e.message));
      const token='a'.repeat(64);let calls=0;
      await page.route('https://hirogari-toroku.github.io/reha1/**',route=>route.fulfill({contentType:'text/html',body:fs.readFileSync(path.join(__dirname,'../index.html'),'utf8')}));
      await page.route('https://static.line-scdn.net/**',route=>route.fulfill({contentType:'application/javascript',body:`window.liff={init:()=>Promise.resolve(),isLoggedIn:()=>true,getProfile:()=>Promise.resolve({userId:'test-liff',displayName:'母'}),getAccessToken:()=> 'fixture-token'};`}));
      await page.route('https://script.google.com/**',route=>{
        calls++;const data=route.request().postDataJSON();
        assert.equal(route.request().method(),'POST');assert.equal(data.accessToken,'fixture-token');
        assert.equal(data.inviteToken,calls===1?token:'');
        return route.fulfill({contentType:'application/json',headers:{'Access-Control-Allow-Origin':'*'},body:JSON.stringify({success:true,userName:'テスト利用者',coupon:{balance:3,status:'OK'},schedules:[{visitDate:'9/20',staffName:'テスト担当',status:'予定'}]})});
      });
      await page.goto('https://hirogari-toroku.github.io/reha1/?mode=user#link='+token);
      await page.getByText('残数 3',{exact:true}).waitFor();
      assert.match(await page.locator('#staff').innerText(),/テスト利用者/);
      assert.equal(await page.evaluate(()=>sessionStorage.getItem('hirogariUserInvite')),null);
      assert.equal(new URL(page.url()).hash,'');
      await page.locator('#scheduleReload').click();
      await page.waitForFunction(()=>document.getElementById('message').innerText==='');
      assert.equal(calls,2);assert.deepEqual(errors,[]);
      assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth),true);
      await page.screenshot({path:`/tmp/hirogari-user-link-${viewport.width}.png`,fullPage:true});
      await page.close();
    }
    console.log('User link browser checks passed on mobile and desktop');
  } finally { await browser.close(); }
})().catch(e=>{console.error(e);process.exitCode=1;});
