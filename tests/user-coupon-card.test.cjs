const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const html=fs.readFileSync(path.join(__dirname,'../index.html'),'utf8');
test('shared staff and user coupon badge retains labels and colors',()=>{
  const context=vm.createContext({document:{createElement:()=>({classList:{add(value){this.value=value;}}})}});
  vm.runInContext(html.slice(html.indexOf('function appendCouponBadge('),html.indexOf('function setResourceData(')),context);
  for(const [coupon,label,color] of [[{balance:2,status:'残数あり'},'残数 2・残数あり','warn'],[{balance:3,status:'OK'},'残数 3',undefined],[{balance:0,status:'OK'},'残数 0','warn'],[{balance:-1,status:'不足'},'残数 -1・不足','danger']]){
    const nodes=[];
    context.appendCouponBadge({appendChild:n=>nodes.push(n)},{coupon});
    assert.equal(nodes[0].textContent,label);
    assert.equal(nodes[0].className,'coupon-pill');
    assert.equal(nodes[0].classList.value,color);
  }
  const nodes=[];
  context.appendCouponBadge({appendChild:n=>nodes.push(n)},{coupon:null});
  assert.equal(nodes.length,0);
});
test('actual user cards and admin preview use the same badge as staff',()=>{
  const user=html.slice(html.indexOf('function renderUserScheduleList()'),html.indexOf('function currentTimeValue()'));
  assert.match(user,/appendCouponBadge\(date, \{ coupon: userCoupon \}\)/);
  const preview=html.slice(html.indexOf('function renderAdminPreview('),html.indexOf('function saveAdminRelationshipRequest('));
  assert.match(preview,/appendCouponBadge\(name, data.targetType === "user" \? \{ coupon: data.coupon \|\| item.coupon \} : item\)/);
  assert.ok(preview.indexOf('appendCouponBadge(summary, { coupon: data.coupon })') < preview.indexOf('if (!schedules.length)'));
});

test('admin user preview returns coupon once even without schedules and never reads chart resources',()=>{
  const c=vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/コード.js'),'utf8'),c);
  c.SpreadsheetApp={getActiveSpreadsheet:()=>({getSheetByName:()=>({})})};
  c.saveLineUserDirectory_=c.ensureScheduleStatusColumns_=()=>{};
  c.isAdminLiffUser_=()=>true;
  c.getAdminUserMap_=()=>({byId:{test:{name:'テスト'}}});
  c.buildVisitStatusIndex_=()=>({});
  c.enrichLiffScheduleItemsWithUserResources_=()=>{throw Error('user preview must not fetch charts');};
  for(const items of [[],[{visitDate:'9/20',userName:'テスト'}]]) {
    let reads=0;
    c.collectActiveSchedulesForUser_=()=>items;
    c.getCouponDisplayMap_=()=>{reads++;return {'テスト':{balance:3,status:'OK'}};};
    const r=c.adminPreviewSchedulesFromLiff_('admin','admin','user','test');
    assert.equal(r.success,true);assert.equal(r.coupon.balance,3);assert.equal(reads,1);
  }
});

test('admin preview renders residual counts on summary and cards, including zero',()=>{
  const nodes=[];
  const element=()=>({children:[],classList:{add(){}},appendChild(n){this.children.push(n);}});
  const root=element();
  const c=vm.createContext({document:{getElementById:()=>root,createElement(){const n=element();nodes.push(n);return n;}}});
  vm.runInContext(html.slice(html.indexOf('function appendCouponBadge('),html.indexOf('function setResourceData(')),c);
  vm.runInContext(html.slice(html.indexOf('function renderAdminPreview('),html.indexOf('function addAdminRelationship(')),c);
  for(const balance of [3,0,-1]) {
    nodes.length=0;
    c.renderAdminPreview({success:true,targetType:'user',targetName:'テスト',coupon:{balance,status:'OK'},schedules:[{visitDate:'9/20',status:'予定'}]});
    const badges=nodes.filter(n=>n.className==='coupon-pill');
    assert.equal(badges.length,2);assert.ok(badges.every(n=>n.textContent==='残数 '+balance));
  }
});
