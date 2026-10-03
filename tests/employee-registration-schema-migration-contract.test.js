const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const filename = 'employee-registration-schema-migration-contract.gs';
const source = fs.readFileSync(path.join(root, filename), 'utf8');
function environment(mutation = '') {
  const forbidden = new Proxy({}, { get() { throw new Error('SIDE_EFFECT'); } });
  const context = vm.createContext({ SpreadsheetApp: forbidden, PropertiesService: forbidden,
    LockService: forbidden, Utilities: forbidden, Logger: forbidden, console: forbidden,
    ScriptApp: forbidden, UrlFetchApp: forbidden, Supabase: forbidden });
  for (const name of ['employee-registration-safety.gs','employee-registration-adapter.gs',
    'employee-registration-operation-ledger.gs','employee-registration-schema-preflight.gs',filename]) {
    vm.runInContext(fs.readFileSync(path.join(root,name),'utf8'), context);
  }
  vm.runInContext(`
    const required = ['employee_id','display_employee_id'].concat(EMPLOYEE_REGISTRATION_INPUT_FIELDS_);
    const headers = Array.from({length:41},(_,i)=>'existing_'+i);
    required.forEach((h,i)=>headers[i]=h); headers[34]='';
    const record={employee_id:'EMP0083',display_employee_id:'W0062',name:'PRIVATE_SENTINEL',
      name_kana:'test',company_code:'MAIN',employment_type:'regular',employment_status:'active',
      hire_date:'2026-04-01',leave_management_target:true,is_driver:false};
    const row=headers.map(h=>record[h]===undefined?'':record[h]);
    let s={contractVersion:1,targetIdentityVerified:true,timeZone:'Asia/Tokyo',
      employees:{values:[headers,row],maxRows:996,maxColumns:41},
      columnInspections:[{column:35,reusable:true,reasons:[]}],
      ledger:{sheetId:null,values:null},emptyLedgerRecovery:null,finalVerification:'UNCONFIRMED',
      externalChecks:{filterViewsReviewed:false,crossSheetFormulasReviewed:false,
        externalReferencesReviewed:false,productionWritesStopped:false,
        backupVerified:false,productionVersionCompatible:false}};
    function operationApplied(){headers[34]='registration_operation_id'}
    function legacyApplied(){headers.push('initial_grant_check_target');row.push('');s.employees.maxColumns=42}
    function ledgerApplied(){s.ledger={sheetId:123,values:[EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.slice()]}}
    function allApplied(){operationApplied();legacyApplied();ledgerApplied()}
    ${mutation}
  `, context);
  return { context, run(code = 'planEmployeeRegistrationSchemaMigration_(s)') {
    return vm.runInContext(code, context);
  } };
}
const plain = value => JSON.parse(JSON.stringify(value));
function result(mutation) { return plain(environment(mutation).run()); }
function manual(mutation) {
  const r = result(mutation);
  assert.equal(r.ok,false); assert.equal(r.state,'MANUAL_INVESTIGATION_REQUIRED');
  assert.deepEqual(r.plan,[]); assert.ok(!Object.hasOwn(r, "productionApplicationAllowed"));
  assert.ok(!JSON.stringify(r).includes('PRIVATE_SENTINEL'));
}
test('first migration has only fixed ordered schema operations',()=>{
  const r=result(''); assert.equal(r.state,'READY_FOR_MIGRATION');
  assert.deepEqual(r.plan,[{type:'ENSURE_COLUMN_CAPACITY',sheet:'employees',minimum:42},
    {type:'SET_HEADER',sheet:'employees',column:35,header:'registration_operation_id'},
    {type:'SET_HEADER',sheet:'employees',column:42,header:'initial_grant_check_target'},
    {type:'CREATE_OPERATION_LEDGER',sheet:'employee_registration_operations',
      headers:['operation_id','input_hash','created_by','status','employee_id','display_employee_id','created_at','updated_at']}]);
  assert.ok(!Object.hasOwn(r, "productionApplicationAllowed")); assert.equal(r.pendingExternalChecks.length,6);
});
for(const [name,setup,absent] of [
  ['AI only','operationApplied()',35],['legacy only','legacyApplied()',42],
  ['ledger only','ledgerApplied()',null]
]) test(name+' excludes applied work',()=>{
  const r=result(setup);assert.equal(r.state,'PARTIAL_MIGRATION');
  assert.ok(!r.plan.some(p=>absent===null?p.type==='CREATE_OPERATION_LEDGER':p.column===absent));
});
test('capacity only is reused without expansion',()=>{
  const r=result("s.employees.maxColumns=42;s.columnInspections.push({column:42,reusable:true,reasons:[]})");
  assert.equal(r.state,'PARTIAL_MIGRATION');assert.ok(!r.plan.some(p=>p.type==='ENSURE_COLUMN_CAPACITY'));
});
test('all applied without final verification is not complete',()=>{
  const r=result('allApplied()');assert.equal(r.state,'VERIFICATION_REQUIRED');assert.deepEqual(r.plan,[]);
});
test('three elements and final verification yield complete',()=>{
  const r=result("allApplied();s.finalVerification='CONFIRMED'");assert.equal(r.state,'COMPLETE');assert.deepEqual(r.plan,[]);
});
test('old preflight already migrated does not imply legacy completion',()=>{
  const r=result('operationApplied();ledgerApplied()');assert.notEqual(r.state,'COMPLETE');
  assert.ok(r.plan.some(p=>p.header==='initial_grant_check_target'));
});
test('prefix ledger completes only missing fixed headers',()=>{
  const r=result('s.ledger={sheetId:123,values:[EMPLOYEE_REGISTRATION_LEDGER_HEADERS_.slice(0,3)]}');
  const p=r.plan.find(p=>p.type==='COMPLETE_LEDGER_HEADERS');assert.equal(p.startColumn,4);
  assert.deepEqual(p.headers,['status','employee_id','display_employee_id','created_at','updated_at']);
});
test('empty ledger without provenance requires investigation',()=>manual('s.ledger={sheetId:123,values:[]}'));
test('empty ledger with matching self-attested provenance still requires review',()=>{
  const r=result('s.ledger={sheetId:123,values:[]};s.emptyLedgerRecovery={sheetId:123,interruptedCreationVerified:true,workRecordVerified:true}');
  assert.equal(r.components.ledger,'EMPTY_UNPROVEN');
  assert.equal(r.state,'MANUAL_INVESTIGATION_REQUIRED');assert.deepEqual(r.plan,[]);
});
test('empty ledger with mismatched identity is rejected',()=>manual('s.ledger={sheetId:123,values:[]};s.emptyLedgerRecovery={sheetId:124,interruptedCreationVerified:true,workRecordVerified:true}'));
test('partly attested empty ledger is rejected',()=>manual('s.ledger={sheetId:123,values:[]};s.emptyLedgerRecovery={sheetId:123,interruptedCreationVerified:true,workRecordVerified:false}'));
for(const [name,mutation] of [
  ['duplicate operation header',"operationApplied();headers[33]='registration_operation_id'"],
  ['wrong operation position',"headers[33]='registration_operation_id'"],
  ['wrong legacy position',"headers[33]='initial_grant_check_target'"],
  ['duplicate legacy header',"legacyApplied();headers[33]='initial_grant_check_target'"],
  ['unsafe AI',"s.columnInspections[0].reusable=false;s.columnInspections[0].reasons=['PROTECTION']"],
  ['AI has existing data',"row[34]='PRIVATE_SENTINEL'"],
  ['unsafe AP',"s.employees.maxColumns=42;s.columnInspections.push({column:42,reusable:false,reasons:['FORMULA']})"],
  ['AP inspection missing',"s.employees.maxColumns=42"],
  ['invalid employee ID',"row[0]='bad'"],
  ['invalid operation value',"operationApplied();row[34]='bad'"],
  ['invalid legacy value',"legacyApplied();row[41]='PRIVATE_SENTINEL'"],
  ['wrong timezone',"s.timeZone='UTC'"],
  ['identity unverified',"s.targetIdentityVerified=false"],
  ['ledger invalid header',"s.ledger={sheetId:123,values:[['PRIVATE_SENTINEL']]}"],
  ['prefix with data',"s.ledger={sheetId:123,values:[['operation_id'],['PRIVATE_SENTINEL']]}"],
  ['unexpected empty header',"headers[33]=''"],
  ['operation with missing ledger',"operationApplied();row[34]='123e4567-e89b-42d3-a456-426614174000'"],
]) test(name+' is manual with no plan',()=>manual(mutation));
for(const [name,mutation] of [
  ['null','s=null'],['missing key','delete s.finalVerification'],['caller plan','s.plan=[]'],
  ['unknown version','s.contractVersion=2'],['bad evidence','s.externalChecks.backupVerified="yes"'],
  ['bad dimension','s.employees.maxColumns=34'],['sparse inspection','s.columnInspections=Array(1)'],
  ['unknown reason',"s.columnInspections[0].reasons=['PRIVATE_SENTINEL']"],
  ['getter',"Object.defineProperty(s,'timeZone',{get(){throw Error('PRIVATE_SENTINEL')},enumerable:true})"],
]) test(name+' rejects malformed snapshot with fixed error',()=>{
  const e=environment(mutation);assert.throws(()=>e.run(),err=>err.message==='EMPLOYEE_REGISTRATION_MIGRATION_SNAPSHOT_INVALID'&&!err.stack.includes('PRIVATE_SENTINEL'));
});
test('legitimate existing ledger operations and flags are preserved',()=>{
  const e=environment(`allApplied();row[34]='123e4567-e89b-42d3-a456-426614174000';row[41]=true;
    s.ledger.values.push([row[34],'v1:'+('a'.repeat(64)),'PRIVATE_SENTINEL','COMPLETED','EMP0083','W0062',new Date(0),new Date(1)]);
    s.finalVerification='CONFIRMED';`);
  const before=e.run('JSON.stringify(s)');const r=plain(e.run());assert.equal(r.state,'COMPLETE');
  assert.deepEqual(r.plan,[]);assert.equal(e.run('JSON.stringify(s)'),before);assert.ok(!JSON.stringify(r).includes('PRIVATE_SENTINEL'));
});
test('same immutable snapshot gives same detached plan',()=>{
  const e=environment(`function freeze(x){if(x&&typeof x==='object'){Object.values(x).forEach(freeze);Object.freeze(x)}}freeze(s)`);
  const a=plain(e.run()),b=plain(e.run());assert.deepEqual(a,b);
});
test('external attestations are diagnostic only and never grant permission',()=>{
  const e=environment('Object.keys(s.externalChecks).forEach(k=>s.externalChecks[k]=true)');
  assert.ok(!Object.hasOwn(e.run(), "productionApplicationAllowed"));
  for(const key of ['filterViewsReviewed','crossSheetFormulasReviewed','externalReferencesReviewed',
    'productionWritesStopped','backupVerified','productionVersionCompatible']) {
    const r=result(`Object.keys(s.externalChecks).forEach(k=>s.externalChecks[k]=true);s.externalChecks.${key}=false`);
    assert.ok(!Object.hasOwn(r, "productionApplicationAllowed"));
  }
});
test('pure source has no service access, logging, networking, or destructive operations',()=>{
  assert.ok(!/\b(?:SpreadsheetApp|PropertiesService|LockService|Utilities|Logger|console|ScriptApp|UrlFetchApp|Supabase)\b/.test(source));
  assert.ok(!/DELETE_ROW|DELETE_COLUMN|DELETE_SHEET|CLEAR|ROLLBACK|REWRITE_EMPLOYEES|GENERATE_LEGACY_UUID|BACKFILL_INITIAL_GRANT/.test(source));
});
test('Date employee cells are validated purely and not mutated',()=>{
  const e=environment("row[headers.indexOf('hire_date')]=new Date('2026-03-31T15:00:00Z')");
  const before=e.run('row[headers.indexOf("hire_date")].getTime()');
  assert.equal(e.run().ok,true);
  assert.equal(e.run('row[headers.indexOf("hire_date")].getTime()'),before);
});
test('claimed final verification cannot complete a missing legacy column',()=>{
  const r=result("operationApplied();ledgerApplied();s.finalVerification='CONFIRMED'");
  assert.notEqual(r.state,'COMPLETE');assert.ok(r.plan.some(p=>p.column===42));
});
test('fully blank ledger grid also requires matched recovery evidence',()=>{
  manual("s.ledger={sheetId:123,values:[['','','']]}");
});
test('complete ledger with malformed operation data is not repaired',()=>{
  manual("ledgerApplied();s.ledger.values.push(Array(8).fill('PRIVATE_SENTINEL'))");
});
test('caller-supplied plans are not part of the input contract',()=>{
  const e=environment("s.plan=[{type:'SET_HEADER',sheet:'employees',column:1,header:'PRIVATE_SENTINEL'}]");
  assert.throws(()=>e.run(),err=>err.message==='EMPLOYEE_REGISTRATION_MIGRATION_SNAPSHOT_INVALID');
});
test('shorter schema cannot create unnamed intermediate columns before AP',()=>{
  manual('headers.length=35;row.length=35;s.employees.maxColumns=35');
});

const safePhysicalAp = "headers.push('');row.push('');s.employees.maxColumns=42;" +
  "s.columnInspections.push({column:42,reusable:true,reasons:[]});";

test('P1 true and false external checks preserve the same structural plan',()=>{
  const unchecked=result('');
  const claimed=result('Object.keys(s.externalChecks).forEach(k=>s.externalChecks[k]=true)');
  assert.deepEqual(claimed.plan,unchecked.plan);assert.equal(claimed.state,unchecked.state);
  assert.equal(unchecked.pendingExternalChecks.length,6);assert.deepEqual(claimed.pendingExternalChecks,[]);
  for(const r of [unchecked,claimed]) assert.ok(!Object.hasOwn(r,'productionApplicationAllowed'));
});
for(const [label,proof] of [
  ['no proof','null'],
  ['matching identity and all recovery flags','{sheetId:123,interruptedCreationVerified:true,workRecordVerified:true}'],
  ['matching identity without verified record','{sheetId:123,interruptedCreationVerified:false,workRecordVerified:false}']
]) test('P1 empty ledger never yields a plan: '+label,()=>{
  const r=result(`s.ledger={sheetId:123,values:[]};s.emptyLedgerRecovery=${proof};
    Object.keys(s.externalChecks).forEach(k=>s.externalChecks[k]=true)`);
  assert.equal(r.state,'MANUAL_INVESTIGATION_REQUIRED');assert.deepEqual(r.plan,[]);
  assert.ok(!Object.hasOwn(r,'productionApplicationAllowed'));
});
test('P2 safe blank AP with AI and ledger applied plans only legacy header',()=>{
  const r=result('operationApplied();ledgerApplied();'+safePhysicalAp);
  assert.equal(r.ok,true);assert.equal(r.state,'PARTIAL_MIGRATION');
  assert.deepEqual(r.plan,[{type:'SET_HEADER',sheet:'employees',column:42,header:'initial_grant_check_target'}]);
  assert.equal(r.components.legacy,'MISSING_SAFE');
});
test('P2 safe blank AI and AP plan both fixed headers without expansion',()=>{
  const r=result('ledgerApplied();'+safePhysicalAp);
  assert.equal(r.ok,true);assert.equal(r.components.operation,'MISSING_SAFE');
  assert.equal(r.components.legacy,'MISSING_SAFE');
  assert.deepEqual(r.plan,[{type:'SET_HEADER',sheet:'employees',column:35,header:'registration_operation_id'},
    {type:'SET_HEADER',sheet:'employees',column:42,header:'initial_grant_check_target'}]);
});
test('P2 only safe AI blank with legacy and ledger applied plans only AI',()=>{
  const r=result('legacyApplied();ledgerApplied()');assert.equal(r.ok,true);
  assert.deepEqual(r.plan,[{type:'SET_HEADER',sheet:'employees',column:35,header:'registration_operation_id'}]);
});
test('P2 capacity-outside AP still needs expansion then AP header only',()=>{
  const r=result('operationApplied();ledgerApplied()');assert.equal(r.ok,true);
  assert.deepEqual(r.plan,[{type:'ENSURE_COLUMN_CAPACITY',sheet:'employees',minimum:42},
    {type:'SET_HEADER',sheet:'employees',column:42,header:'initial_grant_check_target'}]);
});
for(const [label,change] of [
  ['extra unnamed column',"headers[33]=''"],
  ['unsafe AI',"s.columnInspections[0]={column:35,reusable:false,reasons:['PROTECTION']}"],
  ['unsafe AP',"s.columnInspections[1]={column:42,reusable:false,reasons:['NOTE']}"],
  ['AI inspection at wrong index',"s.columnInspections[0].column=36"],
  ['AP inspection at wrong index',"s.columnInspections[1].column=41"],
  ['AI data despite claimed empty inspection',"row[34]='PRIVATE_SENTINEL'"],
  ['AP data despite claimed empty inspection',"row[41]=true"],
  ['duplicate employee IDs',"s.employees.values.push(row.slice())"],
  ['duplicate display IDs',"let second=row.slice();second[0]='EMP0084';s.employees.values.push(second)"],
  ['invalid employee ID',"row[0]='BAD'"],
  ['invalid date',"row[headers.indexOf('hire_date')]='2026-02-30'"],
  ['invalid boolean',"row[headers.indexOf('leave_management_target')]='maybe'"],
  ['invalid employee data',"row[headers.indexOf('name')]=''"],
]) test('P2 virtual headers do not hide '+label,()=>manual(safePhysicalAp+change));
test('P2 virtual headers do not hide invalid applied operation UUID',()=>{
  manual('operationApplied();'+safePhysicalAp+"row[34]='BAD_UUID'");
});
test('P2 virtual headers do not hide invalid existing legacy value',()=>{
  manual("legacyApplied();row[41]='PRIVATE_SENTINEL'");
});
test('P2 validation-only virtual headers leave the original snapshot unchanged',()=>{
  const e=environment(safePhysicalAp);
  const before=e.run('JSON.stringify(s)');const a=plain(e.run());
  assert.equal(a.ok,true);assert.equal(e.run('JSON.stringify(s)'),before);
  assert.equal(e.run('headers[34]'), '');assert.equal(e.run('headers[41]'), '');
  assert.deepEqual(a,plain(e.run()));
});
