const {test}=require('node:test');
const assert=require('node:assert/strict');
const vm=require('node:vm'),fs=require('node:fs'),path=require('node:path');
function fixture(){const logs=[],cache=new Map();const c=vm.createContext({LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},CacheService:{getScriptCache:()=>({get:k=>cache.get(k),put:(k,v)=>cache.set(k,v)})},Utilities:{formatDate:()=> '20260916'},SpreadsheetApp:{getActiveSpreadsheet:()=>({})},saveLiffOperationLog_:(...args)=>logs.push(args)});vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/LiffDiagnostics.js'),'utf8'),c);return {c,logs,cache};}
test('untrusted diagnostics are allowlisted, anonymous and deduplicated',()=>{
 const {c,logs}=fixture();const data={traceId:'test-trace-123',stage:'page_open',accessToken:'secret',userId:'fake',displayName:'fake',message:'secret'};
 assert.equal(c.recordLiffDiagnostic_(data).success,true);c.recordLiffDiagnostic_(data);
 assert.equal(logs.length,1);assert.equal(logs[0][2],'');assert.doesNotMatch(JSON.stringify(logs),/secret|fake/);
 assert.equal(c.recordLiffDiagnostic_({...data,stage:'arbitrary'}).success,false);
 assert.equal(c.recordLiffDiagnostic_({...data,traceId:'bad\nvalue'}).success,false);
});
test('diagnostics fail safely and respect volume limit',()=>{
 const {c,cache,logs}=fixture();cache.set('liffDiagnosticDay:20260916','1000');assert.equal(c.recordLiffDiagnostic_({traceId:'test-trace-123',stage:'page_open'}).success,false);assert.equal(logs.length,0);
 c.saveLiffOperationLog_=()=>{throw Error('storage');};assert.doesNotThrow(()=>c.logUserLinkStage_({traceId:'test-trace-123'},'request_failed',null));
});

test('normal browser stages send nothing, failures send one anonymous diagnostic',()=>{
 const html=fs.readFileSync(path.join(__dirname,'../index.html'),'utf8');
 const calls=[];
 const c=vm.createContext({fetch:(...args)=>{calls.push(args);return Promise.resolve();},GAS_URL:'https://example.test'});
 vm.runInContext(html.slice(html.indexOf('const liffTraceId ='),html.indexOf('    setToday();',html.indexOf('const liffTraceId ='))),c);
 for(const stage of ['page_open','init_ok','profile_ok','user_request','user_pending','user_shown'])c.reportLiffStage(stage);
 assert.equal(calls.length,0);
 c.reportLiffStage('user_network_failed');c.reportLiffStage('user_network_failed');
 assert.equal(calls.length,1);
 assert.deepEqual(Object.keys(JSON.parse(calls[0][1].body)).sort(),['action','stage','traceId']);
});
