import {a1,rect,DATASET_CELL_LIMIT} from './engine.mjs';
import {inspectWorkbookStructure,inspectActiveObjects,selectionAnchor,readContext,supported} from './excel.mjs';
import {aggregationEvidence} from './evidence.mjs';

const statuses=new Set(['passed','failed','unsupported','not_tested']);
const entry=(operation,status,category,detail)=>({operation,status,category,detail});
const failureCategories=new Set(['missing_target','invalid_reference','invalid_range','stale_target','protected_target',
  'unsupported','payload_limit','host_unavailable','hidden_target','partial_read','missing_headers','empty_dataset','excel_sync']);
const suggestions={
  missing_target:'Select the worksheet again.',
  invalid_reference:'Select a valid local range.',
  invalid_range:'Select one contiguous range.',
  stale_target:'Reopen the workbook and retry.',
  protected_target:'Check workbook and worksheet protection.',
  unsupported:'Try an Excel host with the required API.',
  payload_limit:'Select a smaller range.',
  host_unavailable:'Open the add-in inside Excel.',
  hidden_target:'Select a visible worksheet.',
  excel_sync:'Reopen the workbook and retry.',
  timeout:'Retry with a smaller workbook or after Excel responds.',
  review_required:'Run the documented test inside Excel.',
};
const safeCategory=error=>error?.name==='AbortError'?'timeout':failureCategories.has(error?.category)?error.category:'check_failed';

function budget(promise,ms) {
  let timer;
  return Promise.race([promise,new Promise((_,reject)=>{
    timer=setTimeout(()=>reject(new DOMException('Time budget exceeded','AbortError')),ms);
  })]).finally(()=>clearTimeout(timer));
}
const real={inspectWorkbookStructure,inspectActiveObjects,selectionAnchor,readContext,supported};

export async function runLightDiagnostics(office=real,ms=2400) {
  const operations=[];
  try {
    await budget((async()=>{
      const structure=await office.inspectWorkbookStructure();
      if (!structure.active?.id || !structure.sheets.some(s=>s.id===structure.active.id))
        throw Error('Active sheet was not enumerated.');
      operations.push(entry('Workbook and active worksheet','passed','workbook_readable','Excel enumerated the workbook and resolved its active worksheet.'));
      const selection=await office.selectionAnchor();
      const {r,c}=rect(selection.address,DATASET_CELL_LIMIT);
      const cell=await office.readContext({sheetId:selection.id,address:a1(r,c)});
      if (cell.values?.length!==1 || cell.values[0]?.length!==1) throw Error('Cell result was incomplete.');
      operations.push(entry('Single-cell read','passed','cell_readable','Excel returned a cell from the current selection.'));
    })(),ms);
  } catch(error) {
    // Copy results so in-flight readonly calls after timeout cannot mutate the report.
    return [...operations,entry('Workbook access','failed',safeCategory(error),
      suggestions[safeCategory(error)]||'Confirm the workbook is accessible in Excel and retry.')];
  }
  return operations;
}

async function check(operation,work,ms=2500) {
  try {
    const result=await budget(Promise.resolve().then(work),ms);
    const item=entry(operation,result.status??'passed',result.category,result.detail);
    if (typeof result.build==='string' && /^[0-9a-f]{7,40}$/i.test(result.build)) item.build=result.build;
    return item;
  } catch(error) {
    const category=safeCategory(error);
    return entry(operation,category==='unsupported'?'unsupported':'failed',category,
      suggestions[category]||'The check did not complete. Reopen Forma or select a smaller range.');
  }
}

function analyticCheck() {
  const fixture={sheetId:'diagnostic',sheet:'Synthetic',address:'A1:B4',
    values:[['Segment','Amount'],['A',-4],['B',3],['A',10]],errors:[[],[],[],[]]};
  const evidence=aggregationEvidence(fixture,{dimension:'Segment',measure:'Amount'});
  const first=evidence.groups[0];
  if (first.sum!==6||first.records.map(r=>r.measureCell).join(',')!=='B2,B4'||
      evidence.groups.reduce((n,g)=>n+g.sum,0)!==9) throw Error('Synthetic totals did not reconcile.');
  return {category:'synthetic_verified',detail:'A synthetic signed sum reconciled with its exact contributing cells.'};
}

async function backendCheck(request,path,options) {
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),19000);
  try {
    const response=await request(path,{...options,signal:controller.signal});
    if (response.status===401) return {status:'not_tested',category:'authentication_required',detail:'Enter the deployment access password to run this check.'};
    if (response.status===429) return {status:'not_tested',category:'busy',detail:'Try diagnostics again in a minute.'};
    if (!response.ok) return {status:'failed',category:'backend_http',detail:'The service returned HTTP '+response.status+'.'};
    const data=await response.json();
    if (path==='/health') {
      if (data.status!=='ok') throw Error('Unexpected health response.');
      return {category:'backend_reachable',detail:'The hosted health endpoint responded.',build:data.build};
    }
    if (!options?.method) {
      if (typeof data.configured!=='boolean') throw Error('Unexpected configuration response.');
      return {status:data.configured?'passed':'not_tested',category:data.configured?'provider_configured':'provider_unconfigured',
        detail:data.configured?'The server reports that an AI provider key is configured.':'Ask the administrator to configure the AI provider.'};
    }
    if (data.status==='passed'&&data.category==='structured_greeting')
      return {category:'structured_greeting',detail:'Groq returned and Forma validated a synthetic reply.'};
    if (data.status==='not_tested') return {status:'not_tested',category:'provider_unconfigured',detail:'A live Groq provider is not configured.'};
    const allowed=new Set(['provider_refusal','provider_auth','provider_rate_limit','provider_http','network_or_timeout',
      'truncated','invalid_envelope_json','invalid_envelope_shape','invalid_plan_shape','invalid_plan_fields','invalid_content_json',
      'empty_content','invalid_tool_arguments','invalid_tool_sequence','unsupported_finish','explanation_requested_action']);
    const category=allowed.has(data.category)?data.category:'provider_probe_failed';
    return {status:'failed',category,detail:'The synthetic Groq request failed during '+category.replaceAll('_',' ')+'.'};
  } finally {clearTimeout(timer);}
}

export async function runDiagnostics({office=real,request=fetch,probe=false,build='development',host='Excel'}={}) {
  const checks=[];
  for (const version of ['1.1','1.4','1.7','1.8','1.9','1.10','1.13'])
    checks.push(entry('ExcelApi '+version,office.supported(version)?'passed':'unsupported','requirement_set',
      office.supported(version)?'The host reports this requirement set.':'This requirement set is not available in this host.'));
  checks.push(...await runLightDiagnostics(office));
  checks.push(await check('Range, formula and data types',async()=>{
    const selection=await office.selectionAnchor();
    const {r,c}=rect(selection.address,DATASET_CELL_LIMIT);
    const rows=Math.min(2,1048576-r),columns=Math.min(2,16384-c);
    const result=await office.readContext({sheetId:selection.id,address:a1(r,c,rows,columns)});
    for (const matrix of [result.values,result.formulas,result.types,result.formats]) {
      if (!Array.isArray(matrix)||matrix.length!==rows||matrix.some(row=>!Array.isArray(row)||row.length!==columns))
        throw Error('Unexpected Excel range field shape.');
    }
    return {category:'range_fields_read',detail:'Read values, formulas, types and formats from a range of at most four cells. No cell data is included here.'};
  }));
  checks.push(await check('Used range, tables and named objects',async()=>{
    const result=await office.inspectActiveObjects();
    return {category:'objects_read',detail:(result.used?'Populated used range identified. ':'No populated range. ')+
      (result.tableCount===null?'Table enumeration unavailable. ':'Table enumeration succeeded. ')+
      (result.workbookNameCount===null?'Workbook-scoped name enumeration unavailable.':'Workbook-scoped name enumeration succeeded.')};
  }));
  checks.push(await check('Deterministic arithmetic and provenance',analyticCheck));
  checks.push(entry('Preview and write execution','not_tested','review_required',
    'This read-only diagnostic did not execute a preview, apply or undo.'));
  checks.push(entry('Large-range performance','not_tested','review_required',
    'Run the prescribed workbook benchmark on an Excel host.'));
  const network=await Promise.all([
    check('Hosted backend',()=>backendCheck(request,'/health'),5000),
    check('AI configuration',()=>backendCheck(request,'/api/diagnostics'),5000),
    probe?check('Groq connectivity and structured reply',()=>backendCheck(request,'/api/diagnostics',{
      method:'POST',body:'{}',headers:{'Content-Type':'application/json'},
    }),19500):Promise.resolve(entry('Groq connectivity and structured reply','not_tested','probe_not_requested',
      'Enable the optional synthetic provider check to test a live request.')),
  ]);
  return {schema:1,reference:globalThis.crypto?.randomUUID?.()??'diagnostic-'+Date.now(),
    build:network[0].build??(/^[0-9a-f]{7,40}$/i.test(build)?build:'development'),
    host:['Excel','Excel Web','Excel Windows','Excel Mac','Excel iPad'].includes(host)?host:'Excel',
    date:new Date().toISOString(),checks:[...checks,...network].filter(x=>statuses.has(x.status))};
}
