import {a1, datasetColumns, summariseBy} from './engine.mjs';
import {cellRecord, explicitCurrencyCodes, rangeShape} from './workbook-data.mjs';
import {evaluateQueryPlan} from './query-plan.mjs';

export const foldLabel=value=>value.trim().replace(/\s+/gu,' ').toLowerCase();
const numeric=value=>typeof value==='number' && Number.isFinite(value);

export function aggregationEvidence(dataset,fields,{sessionId,normalised=false,queryPlan}={}) {
  const headers=datasetColumns(dataset.values);
  const filter=queryPlan?evaluateQueryPlan(dataset,queryPlan):null;
  if (filter&&(fields.dimension!==filter.fields.dimension||fields.measure!==filter.fields.measure))
    throw Error('The grouped fields do not match the conditional ranking plan.');
  const groupColumn=headers.indexOf(fields.dimension), measureColumn=headers.indexOf(fields.measure);
  if (groupColumn<0 || measureColumn<0 || groupColumn===measureColumn) throw Error('Choose distinct, existing category and measure columns.');
  const bounds=rangeShape(dataset);
  const ids=headers.flatMap((h,i)=>/^(?:order|transaction|record)[ _-]*id$/i.test(h)?[i]:[]);
  const idColumn=ids.length===1?ids[0]:-1;
  const records=[],excluded=[],ledger=new Map(), variants=new Map(), knownIds=new Map();
  let empty=0, repeatedIdRows=0;
  const cell=(r,c)=>a1(bounds.r+r,bounds.c+c);
  for (let i=1;i<bounds.rows;i++) {
    if (filter&&!filter.rows.has(i)) continue;
    const row=dataset.values[i],errors=dataset.errors?.[i]??[];
    if (row.every(v=>v===''||v===null)) {empty++;continue;}
    const label=row[groupColumn], measure=cellRecord(dataset,bounds,i,measureColumn), value=measure.raw;
    const id=idColumn>=0&&!errors.includes(idColumn)?row[idColumn]:null;
    const record={id:dataset.sheetId+'!'+cell(i,groupColumn),sourceRow:bounds.r+i+1,
      label,value,groupCell:cell(i,groupColumn),measureCell:cell(i,measureColumn),
      identifier:id===null||id===''?null:id,identifierCell:idColumn>=0?cell(i,idColumn):null,
      displayed:measure.displayed,format:measure.format,...(filter?{filterCells:filter.rows.get(i)}:{})};
    if (id!==null&&id!=='') {
      if (knownIds.has(id)) repeatedIdRows++;
      knownIds.set(id,true);
    }
    const reason=errors.includes(groupColumn)||errors.includes(measureColumn)?'Excel error':
      label===''||label===null||normalised&&typeof label==='string'&&!foldLabel(label)?'Missing category':
      !['string','number','boolean'].includes(typeof label)||typeof label==='number'&&!numeric(label)?'Invalid category':
      value===''||value===null?'Blank measure':!numeric(value)?'Non-numeric measure':null;
    if (reason) {excluded.push({...record,reason});continue;}
    const group=typeof label==='string'&&normalised?foldLabel(label):label;
    records.push(record);
    if (!ledger.has(group)) ledger.set(group,[]);
    ledger.get(group).push(record);
    if (typeof label==='string') {
      const key=foldLabel(label);
      if (!variants.has(key)) variants.set(key,new Map());
      const names=variants.get(key);
      if (!names.has(label)) names.set(label,[]);
      names.get(label).push(record);
    }
  }
  if (filter&&records.length+excluded.length+empty+filter.unmatched+filter.untyped!==bounds.rows-1)
    throw Error('The filter and aggregation did not account for every source row.');
  const currencyCodes=explicitCurrencyCodes(records.map(r=>r.format));
  if (currencyCodes.length>1)
    throw Error('The measure has more than one currency format. Select a single-currency range or a converted column.');
  // Reuse the existing deterministic summation on just the eligible ledger rows.
  const summary=summariseBy({values:[['Group','Amount'],...records.map(r=>[
    normalised&&typeof r.label==='string'?foldLabel(r.label):r.label,r.value,
  ])]},'Amount','Group');
  if (summary.groups.reduce((n,g)=>n+g.records,0)!==records.length) throw Error('Category totals did not reconcile with their source rows.');
  const groups=summary.groups.map(g=>({label:g.label,sum:g.sum,records:ledger.get(g.label)}));
  const columnRange=c=>a1(bounds.r+1,bounds.c+c,bounds.rows-1,1);
  const source={sheetId:dataset.sheetId,sheet:dataset.sheet,address:dataset.address,table:dataset.table??null,
    groupRange:columnRange(groupColumn),measureRange:columnRange(measureColumn),idRange:idColumn<0?null:columnRange(idColumn)};
  const reasons=Object.create(null);
  for (const item of excluded) reasons[item.reason]=(reasons[item.reason]??0)+1;
  return {version:1,tool:'AggregateData',id:globalThis.crypto?.randomUUID?.()??'aggregate-'+Date.now(),
    sessionId:sessionId??null,createdAt:new Date().toISOString(),source,fields,
    operation:'sum',grouping:normalised?'case-insensitive trimmed whitespace':'exact label',
    observations:{measureFormats:[...new Set(records.map(r=>r.format).filter(Boolean))].slice(0,12),
      groupTypes:[...new Set(records.map(r=>typeof r.label))],measureType:'number'},
    coverage:'full requested range',scope:{includesHiddenAndFilteredRows:true,sampled:false},
    counts:{included:records.length,excluded:excluded.length,empty,repeatedIdRows},reasons,groups,variants,excluded,
    ...(filter?{queryPlan,filter:{condition:filter.condition,unmatched:filter.unmatched,untyped:filter.untyped}}:{}),
    signature:JSON.stringify([source,fields,records,excluded,empty,filter?.signature??null])};
}

export function sameEvidence(previous,next) {
  return previous.version===next.version && previous.grouping===next.grouping && previous.signature===next.signature;
}

export function evidenceIntent(question) {
  const q=question.trim();
  if (/\b(?:fix|normalise|normalize|standardise|standardize)\b.*\b(?:categor(?:y|ies)|labels?)\b/i.test(q) &&
      !/^\s*(?:how|why|what)\b/i.test(q)) return 'normalise';
  if (/^(?:please )?(?:group|calculate|rank|recalculate).+\b(?:case.insensitiv(?:e|ely)|ignore case)\b/i.test(q))
    return 'regroup';
  if (/\blowercase\b|\blisted separately\b|\bcase.sensitive\b/i.test(q)) return 'variants';
  if (/\b(?:where .*come from|(?:which|show).*(?:records|rows).*(?:contribut|behind|total)|(?:source|contributing) (?:rows|records)|those rows|second category)\b/i.test(q)) return 'records';
  return null;
}

export function evidenceGroup(evidence,question,prior) {
  if (/\bsecond category\b/i.test(question)) {
    if (!evidence.groups[1]) throw Error('The calculation has no second category.');
    return evidence.groups[1];
  }
  const match=/\b([0-9][0-9,]*(?:\.\d+)?)\b/.exec(question);
  const value=match?Number(match[1].replaceAll(',','')):null;
  const text=question.toLowerCase();
  const names=evidence.groups.filter(g=>typeof g.label==='string' && text.includes(g.label.toLowerCase()));
  let matches=names.length?names:evidence.groups;
  if (value!==null) matches=matches.filter(g=>g.sum===value);
  if (matches.length>1 && (names.length || value!==null)) {
    const exact=matches.filter(g=>typeof g.label==='string' && question.includes(g.label));
    if (exact.length===1) return exact[0];
    throw Error('More than one category matches. Name the exact category label to see its rows.');
  }
  if ((names.length||value!==null)&&matches.length===0) throw Error('That total does not match the latest calculation. Ask for its category by name.');
  return names.length||value!==null?matches[0]:evidence.groups.find(g=>g.label===prior)??evidence.groups[0];
}
