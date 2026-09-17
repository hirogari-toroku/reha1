const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const code=fs.readFileSync(path.join(__dirname,'../gas/コード.js'),'utf8');
test('LINE log preserves existing five columns and appends display name separately',()=>{
  const c=vm.createContext({});vm.runInContext(code,c);
  const rows=[],ss={getSheetByName:()=>({appendRow:r=>rows.push(Array.from(r))})};
  c.saveLineMessageLog_(ss,'date','受信','登録名','id','1','表示名');
  c.saveLineMessageLog_(ss,'date','送信','管理者通知','id','本文');
  assert.deepEqual(rows[0],['date','受信','登録名','id','1','表示名']);
  assert.equal(rows[1][5],'');
});
test('registration menu records display name with no extra profile fetch or master matching',()=>{
  for(const menu of ['1','4']) {
    const c=vm.createContext({});vm.runInContext(code,c);
    let profiles=0;const logs=[];
    c.SpreadsheetApp={getActiveSpreadsheet:()=>({getSheetByName:()=>({appendRow(){}})})};
    c.getLineDisplayNameFromEvent_=()=>{profiles++;return '表示名';};
    c.saveRegistrationContact_=()=>({success:true});
    c.saveLineMessageLog_=(...args)=>logs.push(args);
    c.getStaffNameFromLineEvent_=()=>{throw Error('unexpected master scan');};
    c.ContentService={createTextOutput:()=>({setMimeType(){return this;}}),MimeType:{TEXT:'text'}};
    c.doPost({postData:{contents:JSON.stringify({events:[{type:'message',source:{userId:'test'},message:{type:'text',text:menu}}]})}});
    assert.equal(profiles,1);assert.equal(logs[0][6],'表示名');
  }
});
