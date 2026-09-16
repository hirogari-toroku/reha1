const {test}=require('node:test'),assert=require('node:assert/strict');
const vm=require('node:vm'),fs=require('node:fs'),path=require('node:path');
function fixture(rows){const c=vm.createContext({});for(const file of ['UserLink.js','AdminUserLinks.js'])vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',file),'utf8'),c);const values=[vm.runInContext('USER_LINK_HEADERS',c),...rows];return c.getAdminUserLinkSummary_({getSheetByName:()=>({getLastRow:()=>values.length,getRange:()=>({getValues:()=>values.slice()})})});}
function row(id,name,relation,state='連携済み',confirmed='確認済み',liff='U'+'1'.repeat(32)){const r=Array(14).fill('');Object.assign(r,{0:id,1:'U'+'2'.repeat(32),2:relation,3:confirmed,4:liff,5:state,12:name});return r;}
test('approved self and family links are summarized without copying IDs',()=>{
 const result=fixture([row('U1','別の表示名','本人'),row('U2','家族表示名','娘','連携済み','確認済み','U'+'3'.repeat(32))]);
 assert.equal(result.U1.liffLinked,true);assert.equal(result.U2.messagingLinked,true);assert.equal(result.U2.viewers[0].relation,'娘');assert.equal(result.U1.viewers[0].name,'別の表示名');assert.doesNotMatch(JSON.stringify(result),/U111111/);
});
test('pending revoked and duplicate LIFF records never appear approved',()=>{
 for(const r of [row('U1','表示名','本人','未連携','未確認'),row('U1','表示名','本人','無効')])assert.equal(fixture([r]).U1.liffLinked,false);
 const r=row('U1','表示名','本人');assert.equal(fixture([r,r.slice()]).U1.liffLinked,false);
 assert.deepEqual(Object.keys(fixture([row('','表示名','')])),[]);
});
