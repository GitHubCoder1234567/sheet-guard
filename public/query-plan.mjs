import {datasetColumns} from './engine.mjs';
import {cellRecord,rangeShape} from './workbook-data.mjs';
import {chooseDiscoveredRegion} from './discovery.mjs';

const quoted=String.raw`"(?:[^"\\\x00-\x1f]|\\["\\])+"`;
const literal=String.raw`(?:${quoted}|""|true|false|-?(?:0|[1-9]\d*)(?:\.\d+)?(?:e[+-]?\d+)?)`;
const fieldToken=String.raw`(?:${quoted}|[^"]+?)`;
const request=new RegExp(String.raw`^(?:rank (${fieldToken}) by (?:the )?total (${fieldToken})|(?:calculate |show (?:me )?)?(?:total|sum of) (${fieldToken}) by (${fieldToken})) where (.+)$`,'i');
const predicate=new RegExp(String.raw`^\s*(${quoted})\s*(<=|>=|!=|=|<|>)\s*(${literal})(?=\s+and\s+|\s*$)`,'i');
const scope=new RegExp(String.raw`\s+in (table|named range|worksheet) (${quoted})$`,'i');
const scopeKeys={table:'table','named range':'namedRange',worksheet:'worksheet'};
const quotedField=new RegExp('^'+quoted+'$');
const operators=new Set(['=','!=','>','<','>=','<=']);
const fieldLimit=255;
const filterLimit=4;
const example='Use exact headers, for example Rank "Team" by total "Hours" where "Status" = "Open" and "Priority" >= 2.';

function fieldPart(value) {
  const trimmed=value.trim();
  if (!trimmed.includes('"')) return trimmed;
  if (!quotedField.test(trimmed)) throw Error(example);
  return JSON.parse(trimmed);
}

export function conditionalQuery(question) {
  const q=question.trim().replace(/^please\s+/i,'').replace(/[.!?]$/,'');
  if (!/^(?:rank |(?:calculate |show (?:me )?)?(?:total|sum of) )/i.test(q)||!/\swhere\s/i.test(q)) return null;
  try {
    const match=request.exec(q);
    if (!match) throw Error(example);
    const [,group,measure,alternateMeasure,alternateGroup,where]=match;
    const filters=[];
    const requestedSource=scope.exec(where);
    const target=requestedSource?{[scopeKeys[requestedSource[1].toLowerCase()]]:JSON.parse(requestedSource[2])}:null;
    let rest=requestedSource?where.slice(0,requestedSource.index):where;
    while (rest) {
      const item=predicate.exec(rest);
      if (!item || filters.length===filterLimit) throw Error(example);
      const [,field,op,value]=item;
      filters.push({field:JSON.parse(field),op,value:/^(true|false)$/i.test(value)?value.toLowerCase()==='true':JSON.parse(value)});
      rest=rest.slice(item[0].length).trim();
      if (rest) rest=rest.replace(/^and\s+/i,'');
    }
    const plan={kind:'conditional-ranking',group:fieldPart(group??alternateGroup),
      measure:fieldPart(measure??alternateMeasure),filters,...(target?{target}:{})};
    validateQueryPlan(plan);
    return {plan};
  } catch {return {error:example};}
}

function exactKeys(object,keys) {
  return object&&typeof object==='object'&&Object.keys(object).length===keys.length&&keys.every(k=>Object.hasOwn(object,k));
}
const fieldName=value=>typeof value==='string'&&value.trim().length>0&&value.length<=fieldLimit;

function validateQueryPlan(plan) {
  const keys=['kind','group','measure','filters'];
  if (plan?.target!==undefined) keys.push('target');
  if (!exactKeys(plan,keys)||plan.kind!=='conditional-ranking'||
      !fieldName(plan.group)||!fieldName(plan.measure)||!Array.isArray(plan.filters)||
      plan.filters.length<1||plan.filters.length>filterLimit)
    throw Error('The conditional ranking plan is invalid.');
  if (keys.includes('target')) {
    const key=Object.values(scopeKeys).find(k=>exactKeys(plan.target,[k]));
    if (!key||!fieldName(plan.target[key])) throw Error('The dataset source is invalid.');
  }
  for (const filter of plan.filters) {
    if (!exactKeys(filter,['field','op','value'])||!fieldName(filter.field)||!operators.has(filter.op))
      throw Error('The filter plan is invalid.');
    const {op,value}=filter;
    if ((typeof value==='number'&&!Number.isFinite(value))||(typeof value==='string'&&value.length>512)||
        !['number','string','boolean'].includes(typeof value)||
        (!['=','!='].includes(op)&&typeof value!=='number'))
      throw Error('Only exact text or Boolean equality and finite numeric comparisons are supported.');
  }
}

const compare={
  '=':(x,y)=>x===y,'!=':(x,y)=>x!==y,'>':(x,y)=>x>y,'<':(x,y)=>x<y,
  '>=':(x,y)=>x>=y,'<=':(x,y)=>x<=y,
};

export function evaluateQueryPlan(dataset,plan) {
  validateQueryPlan(plan);
  if (dataset.partial) throw Error('Conditional ranking requires the full selected range.');
  const bounds=rangeShape(dataset),headers=datasetColumns(dataset.values);
  const resolve=label=>{
    const index=headers.findIndex(h=>h.toLowerCase()===label.trim().toLowerCase());
    if (index<0) throw Error('Column “'+label+'” was not found in the selection. Check the exact header.');
    return index;
  };
  const group=resolve(plan.group),measure=resolve(plan.measure);
  if (group===measure) throw Error('Choose different grouping and measure columns.');
  const predicates=plan.filters.map(({field,op,value})=>{
    const index=resolve(field);
    return {field:headers[index],index,op,value};
  });
  const rows=new Map(),observed=[];
  let unmatched=0,untyped=0;
  for (let r=1;r<bounds.rows;r++) {
    const cells=predicates.map(({index})=>cellRecord(dataset,bounds,r,index));
    observed.push(cells.map(({kind,raw})=>[kind,raw]));
    // Incompatible cells never satisfy a predicate, including an inequality.
    const compatible=cells.every(({raw,kind},i)=>typeof raw===typeof predicates[i].value &&
      ['number','text','empty_text','boolean'].includes(kind));
    if (!compatible) {untyped++;continue;}
    if (!cells.every(({raw},i)=>compare[predicates[i].op](raw,predicates[i].value))) {unmatched++;continue;}
    rows.set(r,cells.map(({address,raw},i)=>({address,value:raw,field:predicates[i].field})));
  }
  return {fields:{dimension:headers[group],measure:headers[measure]},rows,unmatched,untyped,
    condition:predicates.map(({field,op,value})=>JSON.stringify(field)+' '+op+' '+JSON.stringify(value)).join(' and '),
    signature:JSON.stringify(observed)};
}

// Resolve the source before binding columns and evaluating any analytical step.
export async function readPlannedDataset(plan,readDataset) {
  validateQueryPlan(plan);
  const dataset=await readDataset(plan.target,!plan.target,true);
  if (dataset.partial) throw Error('Conditional ranking requires the full requested range.');
  if (!plan.target?.worksheet) return dataset;
  const region=chooseDiscoveredRegion(dataset);
  if (!region) throw Error('Forma could not identify one dataset on the requested worksheet. Use an Excel table, a workbook-scoped named range, or select the header and data.');
  return region;
}
