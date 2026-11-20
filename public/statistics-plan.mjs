import {compileGraph} from './operation-graph.mjs';
import {supportedColumnOperation} from './calculation.mjs';
import {quotedHeader,splitScope,exactKeys,fieldName,planKeys,readDatasetScope} from './query-plan.mjs';

const request=new RegExp(String.raw`^calculate (?:the )?(.+?) of (${quotedHeader})(?: in (?:the )?selection)?$`,'i');
const ratioRequest=new RegExp(String.raw`^calculate (?:the )?ratio of total (${quotedHeader}) to total (${quotedHeader})(?: in (?:the )?selection)?$`,'i');
const percentRequest=new RegExp(String.raw`^calculate total (${quotedHeader}) as a percentage of total (${quotedHeader})(?: in (?:the )?selection)?$`,'i');
const example='Use exact headers, for example Calculate sum and average of "Hours" in table "Schedule".';

function validateStatisticsPlan(plan) {
  if (plan?.kind==='derived-total') {
    if (!exactKeys(plan,planKeys(plan,['kind','operation','numerator','denominator']))||
        !['ratio','percentage'].includes(plan.operation)||!fieldName(plan.numerator)||!fieldName(plan.denominator))
      throw Error('Invalid derived total plan.');
    return;
  }
  if (!exactKeys(plan,planKeys(plan,['kind','column','operations']))||plan.kind!=='column-statistics'||
      !fieldName(plan.column)||!Array.isArray(plan.operations)||plan.operations.length<1||plan.operations.length>4||
      new Set(plan.operations).size!==plan.operations.length||
      !Array.from(plan.operations).every(supportedColumnOperation))
    throw Error('The statistics plan is invalid. Choose up to four different supported operations.');
}

export function statisticsQuery(question) {
  const text=question.trim().replace(/^please\s+/i,'').replace(/[.!?]$/,'');
  const derived=/^calculate (?:(?:the )?ratio of total|total ".+ as a percentage of total)/i.test(text);
  if (!derived&&!/^calculate .+? of "/i.test(text)) return null;
  try {
    const {head,target}=splitScope(text);
    if (derived) {
      const ratio=ratioRequest.exec(head),match=ratio??percentRequest.exec(head);
      if (!match||target&&/ in (?:the )?selection$/i.test(head)) throw Error();
      const plan={kind:'derived-total',operation:ratio?'ratio':'percentage',numerator:JSON.parse(match[1]),
        denominator:JSON.parse(match[2]),...(target?{target}:{})};
      validateStatisticsPlan(plan);
      return {plan};
    }
    const match=request.exec(head);
    if (!match || target&&/ in (?:the )?selection$/i.test(head)) throw Error(example);
    const operations=match[1].toLowerCase().replace(/,\s*and\s+|\s+and\s+/g,',').split(',').map(v=>v.trim());
    const plan={kind:'column-statistics',column:JSON.parse(match[2]),operations,...(target?{target}:{})};
    validateStatisticsPlan(plan);
    return {plan};
  } catch {return {error:derived?'Use two exact headers, for example Calculate the ratio of total "Resolved" to total "Received".':example};}
}

async function snapshotSignature(dataset) {
  if (!globalThis.crypto?.subtle) throw Error('This host cannot verify statistics evidence. Open Forma in a supported Excel host.');
  const identity={sheetId:dataset.sheetId,sheet:dataset.sheet,table:dataset.table??null,address:dataset.address};
  const bytes=new TextEncoder().encode(JSON.stringify([
    identity,dataset.values,dataset.types,dataset.errors,dataset.formats,dataset.display,dataset.formulas,
  ]));
  const signature=await globalThis.crypto.subtle.digest('SHA-256',bytes);
  return {identity,signature:Array.from(new Uint8Array(signature),byte=>byte.toString(16).padStart(2,'0')).join('')};
}

function compilePlan(plan) {
  validateStatisticsPlan(plan);
  if (plan.kind==='derived-total') {
    const nodes=['numerator','denominator'].map(id=>({id,op:'statistic',operation:'sum',column:plan[id]}));
    nodes.push({id:'answer',op:plan.operation,numerator:'numerator',denominator:'denominator'});
    return compileGraph(nodes,['answer']);
  }
  const nodes=plan.operations.map((operation,i)=>({id:'stat'+i,op:'statistic',operation,column:plan.column}));
  return compileGraph(nodes,nodes.map(node=>node.id));
}

export async function executeStatisticsPlan(plan,readDataset) {
  const evaluate=compilePlan(plan);
  const dataset=await readDatasetScope(plan.target,readDataset);
  const results=evaluate(dataset).map(({operation,result,field,counts,observedCurrency,method})=>
    ({operation,result,field,counts,observedCurrency,...(method?{method}:{})}));
  return {plan,results,...await snapshotSignature(dataset)};
}

export async function replayStatisticsEvidence(report,operation,readDataset) {
  const evaluate=compilePlan(report.plan);
  const operations=report.plan.operations??[report.plan.operation];
  if (!operations.includes(operation)) throw Error('That operation is absent from the saved statistics.');
  const {sheetId,table,address}=report.identity;
  const fixed=table?{sheetId,table}:{sheetId,address};
  const dataset=await readDatasetScope(report.plan.target,(target,expand,evidence)=>readDataset(target??fixed,false,evidence));
  const {signature}=await snapshotSignature(dataset);
  if (signature!==report.signature) throw Error('The source snapshot changed. Recalculate the statistics before using its evidence.');
  return evaluate(dataset)[operations.indexOf(operation)];
}
