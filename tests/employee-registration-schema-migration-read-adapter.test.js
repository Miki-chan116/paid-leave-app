const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const filename = 'employee-registration-schema-migration-read-adapter.gs';
function environment(setup = '') {
  const calls = [], logs = [];
  const forbidden = new Proxy({}, {get() {throw new Error('PRIVATE_FORBIDDEN');}});
  const context = vm.createContext({record: name => calls.push(name),
    Logger: {log: value => logs.push(value)}, console: {log: value => logs.push(value), error: value => logs.push(value)},
    PropertiesService: forbidden, ScriptApp: forbidden, UrlFetchApp: forbidden, Supabase: forbidden,
    SpreadsheetApp: forbidden, LockService: forbidden, Utilities: forbidden});
  for (const name of ['employee-registration-safety.gs','employee-registration-adapter.gs',
    'employee-registration-operation-ledger.gs','employee-registration-schema-preflight.gs',
    'employee-registration-schema-migration-contract.gs', filename]) {
    vm.runInContext(fs.readFileSync(path.join(root,name),'utf8'), context);
  }
  assert.equal(calls.length, 0);
  assert.equal(logs.length, 0);
  vm.runInContext(`
    const secret = 'PRIVATE_SENTINEL';
    let failing = '', change = () => {}, opens = 0, time = 10, plannerCalls = 0;
    function api(name, value) { record(name); if (failing === name) throw new Error(secret); return value; }
    const headers = Array.from({length:41}, (_,i) => 'existing_'+i);
    ['employee_id','display_employee_id'].concat(EMPLOYEE_REGISTRATION_INPUT_FIELDS_).forEach((h,i)=>headers[i]=h);
    headers[34]='';
    const employee={employee_id:'EMP0083',display_employee_id:'W0062',name:secret,
      name_kana:'mock',company_code:'MAIN',employment_type:'regular',employment_status:'active',
      hire_date:'2026-04-01',leave_management_target:true,is_driver:false};
    const row=headers.map(h=>employee[h]===undefined?'':employee[h]);
    const e={id:11,name:'employees',values:[headers,row],maxRows:4,maxColumns:41,
      formulas:{},notes:{},validations:{},merged:[],protection:[],sheetProtection:[],named:[],rules:[],filter:null};
    let l=null;
    const config={spreadsheetId:'MOCK_TARGET'};
    const approved={spreadsheetId:'MOCK_TARGET',employeesSheetId:11,ledgerSheetId:null,timeZone:'Asia/Tokyo'};
    let actualId='MOCK_TARGET',zone='Asia/Tokyo',short='',wrongRange=false,missingEmployees=false;
    function range(s,r,c,n,w) {
      return {getRow:()=>api('range.getRow',wrongRange?r+1:r),getColumn:()=>api('range.getColumn',c),
        getNumRows:()=>api('range.getNumRows',n),getNumColumns:()=>api('range.getNumColumns',w),
        getSheet:()=>api('range.getSheet',sheet(s)),
        getValues:()=>api('range.getValues',matrix(s,'values',r,c,n,w)),
        getFormulas:()=>api('range.getFormulas',matrix(s,'formulas',r,c,n,w)),
        getNotes:()=>api('range.getNotes',matrix(s,'notes',r,c,n,w)),
        getDataValidations:()=>api('range.getDataValidations',matrix(s,'validations',r,c,n,w)),
        getMergedRanges:()=>api('range.getMergedRanges',s.merged)};
    }
    function matrix(s,type,r,c,n,w) {
      const rows=Array.from({length:n},(_,i)=>Array.from({length:w},(_,j)=> {
        if(type==='values') return s.values[r+i-1] && c+j-1 < s.values[r+i-1].length ? s.values[r+i-1][c+j-1] : '';
        return s[type][(r+i)+','+(c+j)] ?? (type==='validations'?null:'');
      }));
      if(short===type) rows.pop();
      return rows;
    }
    function sheet(s) {
      return {getName:()=>api('sheet.getName',s.name),getSheetId:()=>api('sheet.getSheetId',s.id),
        getMaxRows:()=>api('sheet.getMaxRows',s.maxRows),getMaxColumns:()=>api('sheet.getMaxColumns',s.maxColumns),
        getLastRow:()=>api('sheet.getLastRow',s.values.length),getLastColumn:()=>api('sheet.getLastColumn',s.values[0]?.length||0),
        getRange:(r,c,n,w)=>api('sheet.getRange',range(s,r,c,n,w)),
        getProtections:type=>api('sheet.getProtections',type===env.spreadsheetApp.ProtectionType.RANGE?s.protection:s.sheetProtection),
        getNamedRanges:()=>api('sheet.getNamedRanges',s.named),
        getConditionalFormatRules:()=>api('sheet.getConditionalFormatRules',s.rules),
        getFilter:()=>api('sheet.getFilter',s.filter)};
    }
    const ss={getId:()=>api('ss.getId',actualId),getSpreadsheetTimeZone:()=>api('ss.getSpreadsheetTimeZone',zone),
      getSheets:()=>api('ss.getSheets',(missingEmployees?[]:[sheet(e)]).concat(l?[sheet(l)]:[])),
      getSheetByName:name=>api('ss.getSheetByName',name==='employees'?(missingEmployees?null:sheet(e)):(l?sheet(l):null))};
    const env={resolveConfiguration:()=>api('config',config),readApprovedTargetRecord:()=>api('approval',approved),
      now:()=>api('now',time++),spreadsheetApp:{ProtectionType:{RANGE:'RANGE',SHEET:'SHEET'},
        DataValidationCriteria:{TEXT_EQUAL_TO:'TEXT_EQUAL_TO',DATE_BETWEEN:'DATE_BETWEEN',VALUE_IN_LIST:'VALUE_IN_LIST',VALUE_IN_RANGE:'VALUE_IN_RANGE',CHECKBOX:'CHECKBOX'},
        openById:id=>{api('openById',id); if(++opens===2) change(); return ss;}}};
    function ledger(values= [EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.slice()]) {
      l={...e,id:12,name:'employee_registration_operations',values:values,maxColumns:8,
        formulas:{},notes:{},validations:{},merged:[],protection:[],sheetProtection:[],named:[],rules:[],filter:null};
      approved.ledgerSheetId=12;
    }
    function operation(){headers[34]='registration_operation_id'}
    function legacy(){headers.push('initial_grant_check_target');row.push('');e.maxColumns=42}
    function mockProtection(type='RANGE', options={}) {
      const defaults={row:1,column:35,rows:1,columns:1,warningOnly:false,domainEdit:false,canEdit:true,
        editors:['mock-editor@example.invalid'],unprotected:[]};
      const details={...defaults,...options};
      return {details:details,
        getProtectionType:()=>api('protection.getProtectionType',type),
        getRange:()=>api('protection.getRange',type==='SHEET'?range(e,1,1,e.maxRows,e.maxColumns):
          range(e,details.row,details.column,details.rows,details.columns)),
        isWarningOnly:()=>api('protection.isWarningOnly',details.warningOnly),
        canDomainEdit:()=>api('protection.canDomainEdit',details.domainEdit),
        canEdit:()=>api('protection.canEdit',details.canEdit),
        getEditors:()=>api('protection.getEditors',details.editors.map(email=>({getEmail:()=>api('user.getEmail',email)}))),
        getUnprotectedRanges:()=>api('protection.getUnprotectedRanges',details.unprotected)};
    }
    function mockValidation(type='TEXT_EQUAL_TO', args=['approved']) {
      return {getCriteriaType:()=>api('validation.getCriteriaType',type),
        getCriteriaValues:()=>api('validation.getCriteriaValues',args),
        getAllowInvalid:()=>api('validation.getAllowInvalid',false),
        getHelpText:()=>api('validation.getHelpText',null)};
    }
    const realPlanner=planEmployeeRegistrationSchemaMigration_;
    planEmployeeRegistrationSchemaMigration_=s=>{plannerCalls++;return realPlanner(s)};
    ${setup}
  `, context);
  return {calls,logs,context,run: code => vm.runInContext(code || 'readEmployeeRegistrationSchemaMigrationSnapshot_(env)',context)};
}
const plain = value => JSON.parse(JSON.stringify(value));
function rejection(setup, beforePlanner = true) {
  const env = environment(setup);
  assert.throws(()=>env.run(), error=> error.message==='EMPLOYEE_REGISTRATION_MIGRATION_READ_FAILED' &&
    !error.message.includes('PRIVATE') && error.cause===undefined);
  if(beforePlanner) assert.equal(env.run('plannerCalls'),0);
  assert.deepEqual(env.logs,[]);
}
test('initial read produces exact private snapshot and safe summary',()=>{
  const env=environment();const r=env.run();
  assert.equal(r.safeSummary.state,'READY_FOR_MIGRATION');
  assert.equal(Object.keys(r.snapshot).length,9);
  assert.equal(r.snapshot.finalVerification,'UNCONFIRMED');
  assert.ok(Object.values(r.snapshot.externalChecks).every(v=>v===false));
  assert.equal(r.snapshot.columnInspections.length,1);
  assert.equal(r.snapshot.columnInspections[0].column,35);
  assert.equal(r.evidence.details.columns[0].coverage.length,9);
  assert.equal(env.run('plannerCalls'),1);
  assert.equal(env.calls.filter(c=>c==='openById').length,2);
  assert.ok(!JSON.stringify(r.safeSummary).includes('PRIVATE'));
  assert.ok(!JSON.stringify(r.safeSummary).includes('MOCK_TARGET'));
  assert.deepEqual(env.logs,[]);
});
for(const [name,setup,state] of [
  ['operation applied','operation()','PARTIAL_MIGRATION'],
  ['legacy applied','legacy()','PARTIAL_MIGRATION'],
  ['capacity only','e.maxColumns=42','PARTIAL_MIGRATION'],
  ['ledger only','ledger()','PARTIAL_MIGRATION'],
  ['all elements applied','operation();legacy();ledger()','VERIFICATION_REQUIRED'],
  ['prefix ledger',"ledger([EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.slice(0,3)])",'PARTIAL_MIGRATION'],
  ['empty ledger','ledger([])','MANUAL_INVESTIGATION_REQUIRED'],
  ['safe two empty targets',"e.maxColumns=42;headers.push('');row.push('');ledger()",'PARTIAL_MIGRATION']]) {
  test(name,()=>{const r=environment(setup).run();assert.equal(r.safeSummary.state,state);
    assert.equal(r.snapshot.emptyLedgerRecovery,null);});
}
for(const [name,setup] of [
 ['configuration target mismatch',"config.spreadsheetId='WRONG'"],
 ['missing approval',"env.readApprovedTargetRecord=()=>null"],
 ['bad approval',"approved.extra=true"],
 ['blank target',"config.spreadsheetId=' '"],
 ['actual target mismatch',"actualId='WRONG'"],
 ['employees sheet ID mismatch','e.id=99'],
 ['ledger sheet ID mismatch','ledger();l.id=99'],
 ['unexpected ledger','ledger();approved.ledgerSheetId=null'],
 ['missing approved ledger','approved.ledgerSheetId=12'],
 ['timezone mismatch',"zone='UTC'"],
 ['employees missing','missingEmployees=true'],
 ['wrong range','wrongRange=true'],
 ['too large physical read','e.maxRows=1000001'],
 ['short values',"short='values'"],
 ['short formulas',"short='formulas'"],
 ['short notes',"short='notes'"],
 ['short validations',"short='validations'"],
 ['undefined metadata','e.named=undefined'],
 ['undefined filter','e.filter=undefined'],
 ['invalid notes',"e.notes['1,35']=42"],
 ['invalid formula type',"e.formulas['1,35']=false"],
 ['ledger blank formula',"ledger();l.formulas['1,1']='=IF(TRUE,\"\",\"\")'"],
 ['header changed',"change=()=>headers[0]='changed'"],
 ['same row count value changed',"change=()=>row[headers.indexOf('name')]='changed'"],
 ['row count changed',"change=()=>e.values.push(row.slice())"],
 ['formula changed',"change=()=>e.formulas['2,1']='=1'"],
 ['ledger changed',"ledger();change=()=>l.values[0][0]='changed'"],
 ['capacity changed','change=()=>e.maxColumns=42'],
 ['timezone changed',"change=()=>zone='UTC'"],
 ['identity changed',"change=()=>actualId='WRONG'"],
 ['column metadata changed',"change=()=>e.notes['1,35']='PRIVATE'"],
 ['column unsafe evidence changed',"e.notes['1,35']='first';change=()=>e.notes['1,35']='second'"],
 ['unread protected metadata','e.protection=undefined'],
 ['invalid date',"row[0]=new Date(NaN)"],
 ['null raw cell','row[0]=null']]) test(name,()=>rejection(setup));
for(const name of ['config','approval','now','openById','ss.getId','ss.getSpreadsheetTimeZone',
 'ss.getSheets','ss.getSheetByName','sheet.getName','sheet.getSheetId','sheet.getMaxRows',
 'sheet.getMaxColumns','sheet.getLastRow','sheet.getLastColumn','sheet.getRange','range.getRow',
 'range.getColumn','range.getNumRows','range.getNumColumns','range.getSheet','range.getValues',
 'range.getFormulas','range.getNotes','range.getDataValidations','range.getMergedRanges',
 'sheet.getProtections','sheet.getNamedRanges','sheet.getConditionalFormatRules','sheet.getFilter']) {
 test('API failure '+name,()=>rejection('failing='+JSON.stringify(name)));
}
for(const [name,setup] of [
 ['unsafe AI value',"row[34]='PRIVATE'"],
 ['unsafe AI formula',"e.formulas['4,35']='=IF(TRUE,\"\",\"\")'"],
 ['unsafe AI note',"e.notes['4,35']='PRIVATE'"],
 ['unsafe AI validation',"e.validations['4,35']={getCriteriaType:()=> 'TEXT_EQUAL_TO',getCriteriaValues:()=>[],getAllowInvalid:()=>false,getHelpText:()=>null}"],
 ['unsafe AI merge','e.merged=[range(e,1,35,1,1)]'],
 ['unsafe AI protection','e.protection=[mockProtection()]'],
 ['unsafe AI sheet protection',"e.sheetProtection=[mockProtection('SHEET')]"],
 ['unsafe AI named range','e.named=[{getRange:()=>range(e,1,35,1,1)}]'],
 ['unsafe AI conditional format','e.rules=[{getRanges:()=>[range(e,1,35,1,1)]}]'],
 ['unsafe AI filter','e.filter={getRange:()=>range(e,1,35,1,1)}'],
 ['unsafe AP','e.maxColumns=42; e.notes["4,42"]="PRIVATE"'],
 ['invalid employee ID',"row[headers.indexOf('employee_id')]=''"],
 ['duplicate employee IDs','e.values.push(row.slice())'],
 ['invalid ledger header',"ledger([['wrong']])"],
 ['invalid ledger row',"ledger();l.values.push(Array(8).fill('PRIVATE'))"],
 ['wrong operation location',"headers[33]='registration_operation_id'"],
 ['wrong legacy data',"operation();legacy();row[41]='PRIVATE'"],
 ['extra blank header',"headers[32]=''"],
 ['invalid operation UUID',"operation();row[34]='PRIVATE';ledger()"]]) {
 test(name,()=>rejection(setup,false));
}
test('Date remains Date; results do not alias service cells',()=>{
 const env=environment("row[headers.indexOf('hire_date')]=new Date('2026-04-01T00:00:00Z')");
 env.run('const readResult=readEmployeeRegistrationSchemaMigrationSnapshot_(env)');
 assert.equal(env.run("readResult.snapshot.employees.values[1][headers.indexOf('hire_date')] instanceof Date"),true);
 env.run("readResult.snapshot.employees.values[1][headers.indexOf('hire_date')].setTime(0)");
 assert.notEqual(env.run("row[headers.indexOf('hire_date')].getTime()"),0);
});
test('source raw arrays remain unchanged',()=>{
 const env=environment();env.run('const before=JSON.stringify(e.values)');env.run();
 assert.equal(env.run('JSON.stringify(e.values)===before'),true);
});
test('empty ledger never generates recovery plan',()=>{
 const env=environment('ledger([])');env.run('const r=readEmployeeRegistrationSchemaMigrationSnapshot_(env)');
 assert.equal(env.run('realPlanner(r.snapshot).plan.length'),0);
});
test('AP outside capacity is never read',()=>{
 const env=environment("const originalRange=sheet; sheet=s=>{const obj=originalRange(s);const get=obj.getRange;obj.getRange=(r,c,n,w)=>{if(c===42)throw Error(secret);return get(r,c,n,w)};return obj}");
 assert.equal(env.run().safeSummary.state,'READY_FOR_MIGRATION');
});
test('private API exception does not reach logs, safe return or error',()=>{
 const env=environment("failing='range.getValues'");
 assert.throws(()=>env.run(),e=>!e.message.includes('PRIVATE') && !e.stack.includes('PRIVATE') && !e.cause);
 assert.deepEqual(env.logs,[]);
});
test('mutation methods absent and top-level service access absent',()=>{
 const source=fs.readFileSync(path.join(root,filename),'utf8');
 assert.ok(!/\.(?:setValues?|insertColumns?|deleteRows?|appendRow|setProperty|createTrigger|newTrigger|fetch)\s*\(/.test(source));
 const env=environment();env.run();assert.deepEqual(env.logs,[]);
 assert.ok(env.calls.every(name=>!name.startsWith('set')));
});
test('valid existing operation rows and native ledger dates are preserved',()=>{
 const env=environment(`operation();legacy();ledger();
   const uuid='123e4567-e89b-42d3-a456-426614174000'; row[34]=uuid;
   l.values.push([uuid,'v1:'+'a'.repeat(64),'MOCK_ADMIN','EMPLOYEE_SAVED',
     'EMP0083','W0062',new Date('2026-04-01T00:00:00Z'),new Date('2026-04-01T01:00:00Z')]);`);
 env.run('const output=readEmployeeRegistrationSchemaMigrationSnapshot_(env)');
 assert.equal(env.run('output.safeSummary.state'),'VERIFICATION_REQUIRED');
 assert.equal(env.run('output.snapshot.ledger.values[1][6] instanceof Date'),true);
 assert.equal(env.run('output.snapshot.ledger.values.length'),2);
 assert.ok(!JSON.stringify(env.run('output.safeSummary')).includes('MOCK_ADMIN'));
 assert.deepEqual(env.logs,[]);
});
test('unsafe and invalid raw data are blocked before planner invocation',()=>{
 for(const setup of ["row[34]='PRIVATE'", "e.notes['4,35']='PRIVATE'",
   "row[headers.indexOf('employee_id')]=''", "ledger([['wrong']])",
   "ledger();l.values.push(Array(8).fill('PRIVATE'))"]) rejection(setup,true);
});
test('out-of-capacity metadata is not treated as inspected empty metadata',()=>{
 rejection('e.named=[{getRange:()=>range(e,1,999,1,1)}]');
});

// P1/P2 regression probes: all failures must stop before planner and any return.
for (const [name, setup] of [
 ['P1 duplicate protection types', "env.spreadsheetApp.ProtectionType.SHEET='RANGE';e.sheetProtection=[mockProtection('SHEET')]"],
 ['P1 missing RANGE', 'delete env.spreadsheetApp.ProtectionType.RANGE'],
 ['P1 missing SHEET', 'delete env.spreadsheetApp.ProtectionType.SHEET'],
 ['P1 unknown RANGE', "env.spreadsheetApp.ProtectionType.RANGE='UNKNOWN'"],
 ['P1 unknown SHEET', "env.spreadsheetApp.ProtectionType.SHEET='UNKNOWN'"],
 ['P1 numeric protection type', 'env.spreadsheetApp.ProtectionType.RANGE=1'],
 ['P1 null protection type', 'env.spreadsheetApp.ProtectionType.SHEET=null'],
 ['P1 swapped protection types', "env.spreadsheetApp.ProtectionType={RANGE:'SHEET',SHEET:'RANGE'}"],
 ['P1 RANGE request fails', "const originalSheet=sheet;sheet=s=>{const obj=originalSheet(s),get=obj.getProtections;obj.getProtections=t=>{if(t==='RANGE')throw Error(secret);return get(t)};return obj}"],
 ['P1 SHEET request fails', "const originalSheet=sheet;sheet=s=>{const obj=originalSheet(s),get=obj.getProtections;obj.getProtections=t=>{if(t==='SHEET')throw Error(secret);return get(t)};return obj}"],
 ['P2 actual protection type mismatch', "operation();legacy();ledger();e.sheetProtection=[mockProtection('RANGE')]"],
 ['P2 sheet warning mode drift', "operation();legacy();ledger();const p=mockProtection('SHEET');e.sheetProtection=[p];change=()=>p.details.warningOnly=true"],
 ['P2 range warning mode drift', "operation();legacy();ledger();const p=mockProtection();e.protection=[p];change=()=>p.details.warningOnly=true"],
 ['P2 range domain edit drift', "operation();legacy();ledger();const p=mockProtection();e.protection=[p];change=()=>p.details.domainEdit=true"],
 ['P2 range current edit capability drift', "operation();legacy();ledger();const p=mockProtection();e.protection=[p];change=()=>p.details.canEdit=false"],
 ['P2 range editor identity drift', "operation();legacy();ledger();const p=mockProtection();e.protection=[p];change=()=>p.details.editors=['other-editor@example.invalid']"],
 ['P2 sheet unprotected ranges drift', "operation();legacy();ledger();const p=mockProtection('SHEET');e.sheetProtection=[p];change=()=>p.details.unprotected=[range(e,2,1,1,1)]"],
 ['P2 invalid warning bool', "operation();legacy();ledger();e.protection=[mockProtection('RANGE',{warningOnly:'true'})]"],
 ['P2 blank editor email', "operation();legacy();ledger();e.protection=[mockProtection('RANGE',{editors:['']})]"],
 ['P2 wrong unprotected sheet', "operation();legacy();ledger();e.sheetProtection=[mockProtection('SHEET',{unprotected:[range(l,1,1,1,1)]})]"],
 ['P2 unprotected range outside capacity', "operation();legacy();ledger();e.sheetProtection=[mockProtection('SHEET',{unprotected:[range(e,1,999,1,1)]})]"],
 ['P2 missing criteria namespace', "operation();legacy();ledger();delete env.spreadsheetApp.DataValidationCriteria;e.validations['4,35']=mockValidation()"],
 ['P2 criteria aliases', "operation();legacy();ledger();env.spreadsheetApp.DataValidationCriteria.TEXT_CONTAINS='TEXT_EQUAL_TO';e.validations['4,35']=mockValidation()"]
]) test(name, () => rejection(setup));

for (const method of ['getProtectionType','getRange','isWarningOnly','canDomainEdit','canEdit','getEditors','getUnprotectedRanges']) {
 test('P2 protection attribute failure '+method, () => rejection(
   "operation();legacy();ledger();e.protection=[mockProtection()];failing='protection."+method+"'"));
}
test('P2 editor email failure',()=>rejection("operation();legacy();ledger();e.protection=[mockProtection()];failing='user.getEmail'"));
for (const method of ['getCriteriaType','getCriteriaValues','getAllowInvalid','getHelpText']) {
 test('P2 validation attribute failure '+method,()=>rejection(
   "operation();legacy();ledger();e.validations['4,35']=mockValidation();failing='validation."+method+"'"));
}
for (const [name, value] of [['undefined','undefined'],['null','null'],['empty',"''"],
 ['unknown',"'PRIVATE_UNKNOWN'"],['number','42'],['boolean','true'],['function','()=>{}'],
 ['unregistered enum object',"{toString:()=> 'TEXT_EQUAL_TO'}"]]) {
 test('P2 invalid criteria type '+name,()=>rejection(
   "operation();legacy();ledger();const invalidRule=mockValidation();invalidRule.getCriteriaType=()=> ("+value+");e.validations['4,35']=invalidRule"));
}
// Default argument would turn undefined into a valid fixture value; test raw API explicitly.
test('P2 getCriteriaType actually undefined',()=>rejection(
 "operation();legacy();ledger();const v=mockValidation();v.getCriteriaType=()=>undefined;e.validations['4,35']=v"));

for (const target of [35,42]) for (const stage of [1,2]) for (const kind of ['header','data','formula']) {
 test('P1 overlap '+target+' '+kind+' pass '+stage,()=>rejection(`
   operation();legacy();ledger();const originalMatrix=matrix;
   matrix=(s,type,r,c,n,w)=>{const values=originalMatrix(s,type,r,c,n,w);
     if(opens===${stage} && s===e && c===${target} && w===1) {
       if(${JSON.stringify(kind)}==='header'&&type==='values')values[0][0]='';
       if(${JSON.stringify(kind)}==='data'&&type==='values')values[1][0]='PRIVATE_MISMATCH';
       if(${JSON.stringify(kind)}==='formula'&&type==='formulas')values[1][0]='=1';
     }return values;};`));
}
for (const target of [35,42]) test('P1 both passes have same internal mismatch '+target,()=>rejection(`
 operation();legacy();ledger();const originalMatrix=matrix;
 matrix=(s,type,r,c,n,w)=>{const values=originalMatrix(s,type,r,c,n,w);
   if(s===e&&c===${target}&&w===1&&type==='values')values[0][0]='';return values;};`));
for (const [name, left, right] of [
 ['Date versus string',"new Date('2026-04-01T00:00:00Z')","'2026-04-01'"],
 ['number versus string','1',"'1'"], ['boolean versus string','true',"'true'"]]) {
 test('P1 typed overlap '+name,()=>rejection(`
  operation();legacy();ledger();row[34]=${left};const originalMatrix=matrix;
  matrix=(s,type,r,c,n,w)=>{const values=originalMatrix(s,type,r,c,n,w);
   if(s===e&&c===35&&w===1&&type==='values')values[1][0]=${right};return values;};`));
}
for(const invalid of ['null','undefined']) test('P1 blank API value is not '+invalid,()=>rejection(`
 operation();legacy();ledger();const originalMatrix=matrix;
 matrix=(s,type,r,c,n,w)=>{const values=originalMatrix(s,type,r,c,n,w);
  if(s===e&&c===35&&w===1&&type==='values')values[1][0]=${invalid};return values;};`));

for (const setup of [
 "e.protection=[mockProtection('RANGE',{column:1})]",
 "operation();legacy();ledger();e.sheetProtection=[mockProtection('SHEET',{warningOnly:true,unprotected:[range(e,2,1,1,1)]})]",
 "operation();legacy();ledger();e.protection=[mockProtection()]"
]) test('normal protection metadata remains valid '+setup.slice(0,35),()=>{
 const f=environment(setup),r=f.run();assert.equal(f.run('plannerCalls'),1);
 const serialized=JSON.stringify([r.safeSummary,r.diagnostics]);
 assert.ok(!serialized.includes('@'));assert.ok(!serialized.includes('PRIVATE'));
 assert.deepEqual(f.logs,[]);
});
test('normal editor and protection enumeration ordering is canonical',()=>{
 const f=environment(`operation();legacy();ledger();
 const a=mockProtection('RANGE',{column:1,editors:['z@example.invalid','a@example.invalid']});
 const b=mockProtection('RANGE',{column:2});e.protection=[a,b];
 change=()=>{e.protection=[b,a];a.details.editors.reverse()};`);
 assert.equal(f.run().evidence.observedEqual,true);
});
for (const [name, type, args] of [
 ['text','TEXT_EQUAL_TO',"['approved']"], ['dates','DATE_BETWEEN',"[new Date('2026-01-01'),new Date('2026-12-31')]"],
 ['list','VALUE_IN_LIST',"[['one','two'],true]"], ['range','VALUE_IN_RANGE','[range(e,1,1,1,1),true]'],
 ['checkbox','CHECKBOX','[]']]) {
 test('normal validation '+name,()=>{
  const f=environment("operation();legacy();ledger();e.validations['4,35']=mockValidation('"+type+"',"+args+")");
  assert.equal(f.run().safeSummary.state,'VERIFICATION_REQUIRED');
 });
}
test('native-like enum identity works without accepting unrelated enum objects',()=>{
 const f=environment(`operation();legacy();ledger();
 const enumType={toString:()=> 'TEXT_EQUAL_TO'};
 env.spreadsheetApp.DataValidationCriteria.TEXT_EQUAL_TO=enumType;
 e.validations['4,35']=mockValidation(enumType);
 env.spreadsheetApp.ProtectionType={RANGE:{toString:()=> 'RANGE'},SHEET:{toString:()=> 'SHEET'}};
 `);
 assert.equal(f.run().safeSummary.state,'VERIFICATION_REQUIRED');
});
test('normal relative-date enum argument is represented by its known name',()=>{
 const f=environment(`operation();legacy();ledger();
 const today={toString:()=> 'TODAY'};env.spreadsheetApp.RelativeDate={TODAY:today};
 env.spreadsheetApp.DataValidationCriteria.DATE_AFTER_RELATIVE='DATE_AFTER_RELATIVE';
 e.validations['4,35']=mockValidation('DATE_AFTER_RELATIVE',[today]);`);
 const r=f.run();assert.equal(r.safeSummary.state,'VERIFICATION_REQUIRED');
 assert.equal(r.evidence.details.columns[0].validations[3][0].arguments[0].relativeDate,'TODAY');
});
test('editor information and private attribute failures never reach output channels',()=>{
 const f=environment("operation();legacy();ledger();e.protection=[mockProtection('RANGE',{editors:['private-editor@example.invalid']})]");
 const r=f.run();assert.ok(JSON.stringify(r.evidence).includes('private-editor@example.invalid'));
 assert.ok(!JSON.stringify([r.safeSummary,r.diagnostics]).includes('private-editor'));
 assert.deepEqual(f.logs,[]);
 rejection("operation();legacy();ledger();e.protection=[mockProtection()];e.protection[0].getEditors=()=>{throw Error('PRIVATE_name_employee_id_uuid_token_email')}");
});
test('protection record processing budget fails closed before planner',()=>rejection(
 "operation();legacy();ledger();e.protection=Array(1001).fill(mockProtection())"));
test('editor evidence processing budget fails closed before planner',()=>rejection(
 "operation();legacy();ledger();e.protection=[mockProtection('RANGE',{editors:Array(1001).fill('mock@example.invalid')})]"));
test('unprotected range evidence processing budget fails closed before planner',()=>rejection(
 "operation();legacy();ledger();e.sheetProtection=[mockProtection('SHEET',{unprotected:Array(1001).fill(range(e,2,1,1,1))})]"));
