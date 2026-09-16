const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const MSG='U'+'1'.repeat(32), LIFF='U'+'2'.repeat(32), OTHER='U'+'3'.repeat(32);
function sheet(rows){const writes=[];return {rows,writes,getLastRow:()=>rows.length,getMaxRows:()=>200,
  getDataRange:()=>({getValues:()=>rows.map(r=>r.slice())}),
  getRange:(r,c,h=1,w=1)=>({getValues:()=>Array.from({length:h},(_,i)=>Array.from({length:w},(_,j)=>rows[r+i-1]?.[c+j-1]??'')),
    setValues(values){writes.push([r,c,values]);values.forEach((row,i)=>{rows[r+i-1]||=[];row.forEach((v,j)=>rows[r+i-1][c+j-1]=v);});},setValue(v){this.setValues([[v]]);}})};}
function fixture(){const c=vm.createContext({});for(const name of ['コード.js','UserLink.js'])vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',name),'utf8'),c);
  const links=sheet([vm.runInContext('USER_LINK_HEADERS.slice()',c),['','','','未確認',LIFF,'未連携','old-hash','old-expiry','','','old-sent','','母の表示名','保持するメモ']]);
  const master=sheet([['利用者ID','利用者名','状態'],['U002','山田太郎','利用中']]);
  const directory=sheet([Array(13).fill(''),['','',MSG,'','母の表示名','','','','','','','山田太郎 母','']]);
  const ss={getSheetByName:n=>n==='利用者LINE連携'?links:n==='利用者マスタ'?master:n==='LINEユーザー一覧'?directory:sheet([[]])};
  c.SpreadsheetApp={getActiveSpreadsheet:()=>ss,flush:()=>{}};c.Session={getActiveUser:()=>({getEmail:()=> 'editor@example.test'})};
  c.getStaffNameCached_=()=> '未登録';c.LockService={getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})};
  const input={userId:'U002',messagingId:MSG,liffId:LIFF,relation:'母',confirmed:true};return {c,ss,links,master,directory,input};}

test('confirmed mother links captured IDs without modifying masters or historical columns',()=>{
  const f=fixture();f.c.UrlFetchApp={fetch:()=>{throw Error('must not send');}};f.c.userLinkConfirmAccount(f.input);
  assert.equal(f.links.rows[1][0],'U002');assert.equal(f.links.rows[1][1],MSG);assert.equal(f.links.rows[1][2],'母');assert.equal(f.links.rows[1][4],LIFF);assert.equal(f.links.rows[1][5],'連携済み');
  assert.equal(f.links.rows[1][6],'old-hash');assert.equal(f.links.rows[1][7],'old-expiry');assert.equal(f.links.rows[1][10],'old-sent');assert.equal(f.links.rows[1][13],'保持するメモ');assert.equal(f.master.writes.length,0);assert.equal(f.c.userLinkResolve_(f.ss,LIFF).id,'U002');
});
test('capture appends only once and never authorizes by display name',()=>{
  const f=fixture();f.c.userLinkCapturePending_(f.ss,{userId:OTHER,displayName:'山田太郎'});const count=f.links.rows.length;
  f.c.userLinkCapturePending_(f.ss,{userId:OTHER,displayName:'山田太郎'});assert.equal(f.links.rows.length,count);assert.equal(f.c.userLinkResolve_(f.ss,OTHER),null);assert.equal(f.links.rows.at(-1)[0],'');assert.equal(f.master.writes.length,0);
});
test('unknown, ambiguous, unconfirmed or already-used identities are rejected',()=>{
  for(const mutate of [f=>f.input.confirmed=false,f=>f.input.liffId=OTHER,f=>f.input.messagingId=OTHER,f=>f.links.rows.push(f.links.rows[1].slice()),f=>f.links.rows[1][5]='無効',f=>f.master.rows.push(['U003','山田太郎','利用中']),f=>f.links.rows.push(['U003',MSG])]){
    const f=fixture();mutate(f);assert.throws(()=>f.c.userLinkConfirmAccount(f.input));assert.equal(f.links.writes.length,0);
  }
  const f=fixture();f.c.userLinkConfirmAccount(f.input);assert.throws(()=>f.c.userLinkConfirmAccount(f.input));
});
test('revoked records never become pending again',()=>{
  const f=fixture();f.c.userLinkConfirmAccount(f.input);f.links.rows[1][5]='無効';assert.throws(()=>f.c.userLinkResolve_(f.ss,LIFF));const n=f.links.rows.length;f.c.userLinkCapturePending_(f.ss,{userId:LIFF,displayName:'母'});assert.equal(f.links.rows.length,n);assert.equal(f.links.rows[1][5],'無効');
});
test('unlinked requests capture verified profile only and ignore old invitation token',()=>{
  const f=fixture();f.c.userLinkVerify_=()=>({userId:OTHER,displayName:'利用者家族'});const r=f.c.userLinkRequest_({accessToken:'verified',inviteToken:'old-token',lineUserId:LIFF});assert.equal(r.pending,true);assert.equal(r.schedules.length,0);assert.equal(f.links.rows.at(-1)[4],OTHER);assert.equal(f.links.rows.at(-1)[0],'');
});
test('staff do not enter pending queue and retain original initialization',()=>{
  const f=fixture();f.c.userLinkVerify_=()=>({userId:LIFF});f.c.getStaffNameCached_=()=> '担当';assert.equal(f.c.userLinkRequest_({accessToken:'verified'}).success,false);assert.equal(f.links.writes.length,0);
  f.c.getLiffInitDataFromDisplayMaster_=()=>({staffName:'担当'});f.c.initLiffApp_=()=>({success:true,staffName:'担当'});f.c.liffResponse_=(_,d)=>d;assert.equal(f.c.doGet({parameter:{action:'commonInit',lineUserId:LIFF}}).staffName,'担当');
});
test('common entrance routes unknown IDs without notifying as unregistered staff',()=>{
  const f=fixture();f.c.getLiffInitDataFromDisplayMaster_=()=>null;f.c.liffResponse_=(_,d)=>d;f.c.initLiffApp_=()=>{throw Error('must not log as staff');};assert.equal(f.c.doGet({parameter:{action:'commonInit',lineUserId:OTHER}}).role,'userOrPending');
});
test('LINE verification requires correct channel and positive expiry',()=>{
  for(const check of [{code:401,body:{}},{code:200,body:{client_id:'other',expires_in:10}},{code:200,body:{client_id:'2010856600',expires_in:0}}]){
    const f=fixture();let n=0;f.c.UrlFetchApp={fetch:()=>{n++;return {getResponseCode:()=>check.code,getContentText:()=>JSON.stringify(check.body)};}};assert.throws(()=>f.c.userLinkVerify_('token'));assert.equal(n,1);
  }
  const f=fixture();let n=0;f.c.UrlFetchApp={fetch:()=>({getResponseCode:()=>200,getContentText:()=>JSON.stringify(++n===1?{client_id:'2010856600',expires_in:100}:{userId:LIFF})})};assert.equal(f.c.userLinkVerify_('token').userId,LIFF);
});
test('network exceptions never expose tokens or write pending records',()=>{
  const f=fixture();f.c.UrlFetchApp={fetch:()=>{throw Error('access_token=secret');}};assert.doesNotMatch(f.c.userLinkRequest_({accessToken:'token'}).message,/secret/);assert.equal(f.links.writes.length,0);
});
test('read-only response omits charts and mutation IDs and preserves completed visit text',()=>{
  const f=fixture();f.c.userLinkConfirmAccount(f.input);f.c.userLinkVerify_=()=>({userId:LIFF});f.c.buildVisitStatusIndex_=()=>({});f.c.collectActiveSchedulesForUser_=(_,name)=>{assert.equal(name,'山田太郎');return [{visitDate:'9/20',staffName:'担当',chartUrl:'private',scheduleId:'mutation-id',lastVisitText:'実績'}];};f.c.getCouponDisplayMap_=()=>({'山田太郎':{balance:3}});const r=f.c.userLinkRequest_({accessToken:'valid'});assert.equal(r.success,true);assert.equal(r.coupon.balance,3);assert.equal(r.schedules[0].chartUrl,undefined);assert.equal(r.schedules[0].scheduleId,undefined);assert.equal(r.schedules[0].lastVisitText,'実績');assert.equal(f.master.writes.length,0);
});
test('legacy unsigned endpoint returns no private data',()=>{const f=fixture();f.c.liffResponse_=(_,d)=>d;assert.equal(f.c.doGet({parameter:{action:'getUserSchedules',displayName:'山田太郎'}}).success,false);});
