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
  assert.match(preview,/if \(data.targetType === "staff"\) \{\s*appendTitleChartLink\(name, item\);\s*\}\s*appendCouponBadge\(name, item\);/);
});
