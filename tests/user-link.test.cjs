const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const path = require('node:path');
const ID = 'U' + '1'.repeat(32), OTHER = 'U' + '2'.repeat(32);
const TOKEN = 'a'.repeat(64);
const digest = value => crypto.createHash('sha256').update(value).digest('hex');

function sheet(rows) {
  const writes = [];
  return { rows, writes, getLastRow: () => rows.length, getMaxRows: () => 200,
    getDataRange: () => ({getValues: () => rows.map(r => r.slice())}),
    getRange: (r, c, h = 1, w = 1) => ({
      getValues: () => Array.from({length:h}, (_, i) => Array.from({length:w}, (_, j) => rows[r+i-1]?.[c+j-1] ?? '')),
      setValues(values) { writes.push([r,c,values]); values.forEach((row,i) => { rows[r+i-1] ||= []; row.forEach((v,j) => rows[r+i-1][c+j-1] = v); }); },
      setValue(value) { this.setValues([[value]]); }
    }) };
}

function fixture() {
  const c = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/コード.js'),'utf8'),c);
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/UserLink.js'),'utf8'),c);
  const headers = vm.runInContext('USER_LINK_HEADERS.slice()',c);
  const row = ['U002',ID,'母','確認済み','','招待中',digest(TOKEN),new Date(Date.now()+60000),'','admin','','山田太郎','母の表示名','保持するメモ'];
  const links = sheet([headers,row]);
  const master = sheet([['利用者ID','利用者名','状態'],['U002','山田太郎','利用中']]);
  const directory = sheet([Array(13).fill(''),['','',ID,'','母の表示名','','','','','','','山田太郎 母','']]);
  const ss = {getSheetByName: name => name === '利用者LINE連携' ? links : name === '利用者マスタ' ? master : name === 'LINEユーザー一覧' ? directory : sheet([[]])};
  c.SpreadsheetApp = {getActiveSpreadsheet:()=>ss,flush:()=>{}};
  c.Utilities = {DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(_,s)=>[...crypto.createHash('sha256').update(s).digest()],getUuid:()=>crypto.randomUUID()};
  c.Session = {getActiveUser:()=>({getEmail:()=> 'editor@example.test'})};
  let releases=0;
  c.LockService={getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{releases++;}})};
  return {c,ss,links,master,directory,releases:()=>releases};
}

test('mother links to user without changing master, notes or send timestamp',()=>{
  const f=fixture(), before=JSON.stringify(f.master.rows);
  f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER});
  assert.equal(f.links.rows[1][4],OTHER);assert.equal(f.links.rows[1][5],'連携済み');
  assert.equal(f.links.rows[1][2],'母');assert.equal(f.links.rows[1][13],'保持するメモ');
  assert.equal(f.links.rows[1][10],'');assert.equal(JSON.stringify(f.master.rows),before);
  assert.equal(f.c.userLinkResolve_(f.ss,OTHER).id,'U002');
  assert.equal(f.releases(),1);
});

test('same-identity retry is idempotent; another identity cannot reuse',()=>{
  const f=fixture();f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER});
  const n=f.links.writes.length;f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER});
  assert.equal(f.links.writes.length,n);
  assert.throws(()=>f.c.userLinkClaim_(f.ss,TOKEN,{userId:ID}),/使用済み/);
});

test('expired, unapproved, disabled and tampered invitations do not write',()=>{
  for(const mutate of [f=>f.links.rows[1][7]=new Date(0),f=>f.links.rows[1][3]='未確認',f=>f.links.rows[1][5]='無効',f=>f.links.rows[1][6]='wrong']) {
    const f=fixture();mutate(f);assert.throws(()=>f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER}));assert.equal(f.links.writes.length,0);
  }
});

test('ambiguous user names, IDs and existing identity mappings fail closed',()=>{
  for(const mutate of [f=>f.master.rows.push(['U003','山田太郎','利用中']),f=>f.master.rows.push(['U002','別人','利用中']),f=>f.links.rows.push(['U003',OTHER,'本人','確認済み',OTHER,'連携済み'])]) {
    const f=fixture();mutate(f);assert.throws(()=>f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER}));assert.equal(f.links.writes.length,0);
  }
});

test('revocation blocks the next read and display names do not authorize',()=>{
  const f=fixture();assert.equal(f.c.userLinkResolve_(f.ss,'母の表示名'),null);
  f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER});f.links.rows[1][5]='無効';
  assert.throws(()=>f.c.userLinkResolve_(f.ss,OTHER),/連携状態/);
});

test('invitation creation keeps notes and does not send any message',()=>{
  const f=fixture();f.c.UrlFetchApp={fetch:()=>{throw Error('no network or sending allowed');}};
  const response=f.c.userLinkCreateInvitation({userId:'U002',messagingId:ID,relation:'母',confirmed:true});
  assert.match(response.message,/#link=[a-f0-9]{64}/);
  assert.equal(f.links.rows[1][13],'保持するメモ');assert.equal(f.links.rows[1][10],'');
  assert.equal(f.master.writes.length,0);
  assert.notEqual(f.links.rows[1][6],digest(TOKEN));
  assert.throws(()=>f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER}),/無効/);
});

test('missing confirmation and unknown official LINE recipient are rejected',()=>{
  const f=fixture();
  assert.throws(()=>f.c.userLinkCreateInvitation({confirmed:false}));
  assert.throws(()=>f.c.userLinkCreateInvitation({confirmed:true,relation:'母',messagingId:OTHER,userId:'U002'}));
  assert.equal(f.links.writes.length,0);
});

test('LINE token must be valid, unexpired and issued to this channel',()=>{
  for(const check of [{code:401,body:{}},{code:200,body:{client_id:'other',expires_in:100}},{code:200,body:{client_id:'2010856600',expires_in:0}}]) {
    const f=fixture();let calls=0;f.c.UrlFetchApp={fetch:()=>{calls++;return {getResponseCode:()=>check.code,getContentText:()=>JSON.stringify(check.body)};}};
    assert.throws(()=>f.c.userLinkVerify_('token'));assert.equal(calls,1);
  }
  const f=fixture();let calls=0;f.c.UrlFetchApp={fetch:()=>{calls++;return {getResponseCode:()=>200,getContentText:()=>JSON.stringify(calls===1?{client_id:'2010856600',expires_in:100}:{userId:OTHER,displayName:'母'})};}};
  assert.equal(f.c.userLinkVerify_('token').userId,OTHER);
});

test('raw network exceptions never expose token URLs',()=>{
  const f=fixture();f.c.UrlFetchApp={fetch:()=>{throw Error('https://line.example/?access_token=secret');}};
  assert.doesNotMatch(f.c.userLinkRequest_({accessToken:'token'}).message,/secret/);
});

test('authorized response contains only this user schedule and coupon, no chart URL',()=>{
  const f=fixture();f.c.userLinkClaim_(f.ss,TOKEN,{userId:OTHER});
  f.c.userLinkVerify_=()=>({userId:OTHER});f.c.buildVisitStatusIndex_=()=>({});
  f.c.collectActiveSchedulesForUser_=(_,name)=>{assert.equal(name,'山田太郎');return [{visitDate:'9/20',staffName:'担当',chartUrl:'private',userName:'山田太郎'}];};
  f.c.getCouponDisplayMap_=()=>({'山田太郎':{balance:3},別人:{balance:9}});
  const result=f.c.userLinkRequest_({accessToken:'token'});
  assert.equal(result.success,true);assert.equal(result.coupon.balance,3);assert.equal(result.schedules[0].chartUrl,undefined);
  assert.equal(f.master.writes.length,0);
});

test('legacy unsigned user endpoint returns no private schedules; staff endpoint unchanged',()=>{
  const f=fixture();f.c.liffResponse_=(_,data)=>data;
  assert.equal(f.c.doGet({parameter:{action:'getUserSchedules',displayName:'山田太郎'}}).success,false);
  f.c.getSchedulesForLiff_=id=>({success:true,staff:id});
  assert.equal(f.c.doGet({parameter:{action:'getSchedules',lineUserId:'staff'}}).staff,'staff');
});
