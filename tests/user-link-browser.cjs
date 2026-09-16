const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');

(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'chrome' });
  try {
    for (const viewport of [{width:390,height:844},{width:1440,height:900}]) {
      for (const role of ['user', 'pending', 'staff']) {
      const page = await browser.newPage({viewport});
      const errors=[];page.on('pageerror',e=>errors.push(e.message));
      let calls=0;
      await page.route('https://hirogari-toroku.github.io/reha1/**',route=>route.fulfill({contentType:'text/html',body:fs.readFileSync(path.join(__dirname,'../index.html'),'utf8')}));
      await page.route('https://static.line-scdn.net/**',route=>route.fulfill({contentType:'application/javascript',body:`window.liff={init:()=>Promise.resolve(),isLoggedIn:()=>true,getProfile:()=>Promise.resolve({userId:'test-liff',displayName:'母'}),getAccessToken:()=> 'fixture-token'};`}));
      await page.route('https://script.google.com/**',route=>{
        const url=new URL(route.request().url());
        if (route.request().method()==='GET') {
          const action=url.searchParams.get('action');
          assert.ok(['commonInit','getSchedules'].includes(action));
          const data=role==='staff'?{success:true,staffName:'テスト担当',users:[{name:'テスト利用者'}],schedules:[{scheduleId:'test',userName:'テスト利用者',visitDate:'9/20',visitDateValue:'2026-09-20',status:'予定',chartUrl:'https://example.test/chart'}]}:{success:false,role:'userOrPending'};
          return route.fulfill({contentType:'application/javascript',body:url.searchParams.get('callback')+'('+JSON.stringify(data)+')'});
        }
        const data=route.request().postDataJSON();
        if(data.action==='liffDiagnostic') {
          assert.deepEqual(Object.keys(data).sort(),['action','stage','traceId']);
          return route.fulfill({contentType:'application/json',body:'{"success":true}'});
        }
        calls++;
        assert.equal(route.request().method(),'POST');assert.equal(data.accessToken,'fixture-token');
        assert.equal(data.inviteToken,undefined);
        return route.fulfill({contentType:'application/json',headers:{'Access-Control-Allow-Origin':'*'},body:JSON.stringify(role==='pending'?{success:false,pending:true,schedules:[],message:'管理者が確認・登録後に予定を表示します。'}:{success:true,userName:'テスト利用者',coupon:{balance:3,status:'OK'},schedules:[{visitDate:'9/20',staffName:'テスト担当',status:'完了',lastVisitText:'開始 18:00 / 終了 19:00'}]})});
      });
      await page.goto('https://hirogari-toroku.github.io/reha1/');
      if (role==='user') {
        await page.locator('.schedule-card-title .coupon-pill').getByText('残数 3',{exact:true}).waitFor();
        assert.match(await page.locator('#staff').innerText(),/テスト利用者/);
        await page.getByText('開始 18:00 / 終了 19:00',{exact:true}).waitFor();
        await page.locator('#scheduleReload').click();
        await page.waitForFunction(()=>document.getElementById('message').innerText==='');
        assert.equal(calls,2);
      } else if(role==='pending') {
        await page.getByText('登録確認待ち',{exact:true}).waitFor();assert.equal(calls,1);
        assert.doesNotMatch(await page.locator('#scheduleList').innerText(),/テスト利用者/);
      } else {
        await page.getByRole('button',{name:'キャンセル',exact:true}).waitFor();assert.equal(calls,0);
        await page.getByRole('link',{name:'カルテ',exact:true}).waitFor();
      }
      if(role!=='staff') {
        assert.equal(await page.getByRole('button',{name:/^(開始|終了|変更|キャンセル)$/}).count(),0);
        assert.equal(await page.getByRole('link',{name:'カルテ',exact:true}).count(),0);
      }
      assert.deepEqual(errors,[]);
      assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth),true);
      await page.screenshot({path:`/tmp/hirogari-common-${role}-${viewport.width}.png`,fullPage:true});
      await page.close();
      }
    }
    console.log('Common entrance user, pending and staff checks passed on mobile and desktop');
  } finally { await browser.close(); }
})().catch(e=>{console.error(e);process.exitCode=1;});
