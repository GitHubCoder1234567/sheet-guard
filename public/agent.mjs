import {a1, datasetColumns, formula, formulaReferences, profile, rect, summariseBy, DATASET_CELL_LIMIT} from './engine.mjs';

const tools=new Set(['InspectWorkbook','InspectWorksheet','InspectRange','AnalyseColumn','AggregateData',
  'FindDuplicates','AnalyseFormulas','CompareDatasets','GenerateChartProposal','GenerateFormattingProposal',
  'GenerateDashboardProposal','ExplainFormula']);
const operations=new Set(['count','total','highest','lowest']);
const shorten=v=>typeof v==='string'?v.slice(0,180):v;
const source=d=>d.sheet+'!'+d.address;

export function toolTarget(call,sheets,selection) {
  if (!tools.has(call.tool)) throw Error('Forma could not use the requested workbook tool.');
  for (const name of ['sheet','address','measure','dimension','operation'])
    if (typeof call[name]!=='string' || call[name].length>120) throw Error('The workbook tool arguments were invalid.');
  const sheet=call.sheet.trim()?sheets.find(s=>s.name.toLowerCase()===call.sheet.trim().toLowerCase()):undefined;
  if (call.sheet.trim() && (!sheet || sheet.visibility!=='Visible'))
    throw Error('Choose a visible worksheet for this request.');
  if (!call.address) {
    if (sheet && selection?.id!==sheet.id && !['InspectWorksheet','InspectWorkbook'].includes(call.tool))
      throw Error('Select the dataset on '+sheet.name+' or specify its range.');
    return sheet?{sheetId:sheet.id}:null;
  }
  rect(call.address,call.tool==='InspectRange'?2400:DATASET_CELL_LIMIT);
  const sheetId=sheet?.id??selection?.id;
  if (!sheetId) throw Error('Select a worksheet first.');
  return {sheetId,address:call.address};
}

function groupedReply(call,d) {
  if (!operations.has(call.operation)) throw Error('Choose a count, total, highest or lowest calculation.');
  if (call.operation==='count') {
    datasetColumns(d.values);
    const count=d.values.slice(1).filter(row=>row.some(v=>v!=='' && v!==null)).length;
    return count.toLocaleString()+' records in '+source(d)+', excluding the header and empty rows.';
  }
  if (!call.measure) throw Error('Which numeric column should Forma total?');
  if (!call.dimension && ['highest','lowest'].includes(call.operation))
    throw Error('Which column should Forma use to group the totals?');
  let result;
  if (!call.dimension) {
    const columns=datasetColumns(d.values);
    let header='Forma total';
    while (columns.some(c=>c.toLowerCase()===header.toLowerCase())) header+='_';
    result=summariseBy({values:d.values.map((row,i)=>[...row,i?'All records':header]),errors:d.errors},call.measure,header);
  } else result=summariseBy(d,call.measure,call.dimension);
  if (!result.groups.length) return 'No eligible numeric records in '+source(d)+'.';
  const fmt=n=>n.toLocaleString(undefined,{maximumFractionDigits:10});
  let lines;
  if (call.operation==='total') lines=result.groups.slice(0,12).map(g=>g.label+': '+fmt(g.sum));
  else {
    const best=call.operation==='highest'?result.groups[0].sum:result.groups.at(-1).sum;
    const ties=result.groups.filter(g=>g.sum===best);
    const winners=ties.slice(0,8).map(g=>g.label).join(', ');
    lines=[(call.operation==='highest'?'Highest':'Lowest')+' total '+result.measure+': '+winners+' · '+fmt(best)+
      (ties.length>8?' ('+ties.length+' tied groups)':'')];
  }
  const scope=result.count+' rows included, '+result.omitted+' with missing or invalid measure/group values omitted.';
  return [lines.join('\n'),'Source: '+source(d)+'. '+scope+
    (call.operation==='total' && result.groups.length>12?' Showing the 12 largest groups.':'')].join('\n\n');
}

function duplicates(call,d) {
  const columns=datasetColumns(d.values);
  const key=call.measure?columns.findIndex(h=>h.toLowerCase()===call.measure.trim().toLowerCase()):-1;
  if (call.measure && key<0) throw Error('Column "'+call.measure+'" was not found.');
  const start=rect(d.address,DATASET_CELL_LIMIT);
  const seen=new Map(), groups=new Map();
  let omitted=0, repeats=0;
  for (let i=1;i<d.values.length;i++) {
    const row=d.values[i];
    if (!row.some(v=>v!=='' && v!==null)) continue;
    if (key>=0 ? d.errors[i]?.includes(key)||row[key]===''||row[key]===null : d.errors[i]?.length) {omitted++;continue;}
    const match=key<0?JSON.stringify(row):row[key];
    if (!seen.has(match)) {seen.set(match,i);continue;}
    repeats++;
    if (!groups.has(match)) groups.set(match,[seen.get(match)]);
    if (groups.get(match).length<5) groups.get(match).push(i);
  }
  const examples=[...groups.values()].slice(0,6).map(indices=>indices.map(i=>String(start.r+i+1)).join(', '));
  return (key<0?'Exact matching rows':'Repeated keys in '+columns[key])+': '+groups.size+' groups, '+repeats+
    ' additional matching records. Source: '+source(d)+'. '+omitted+' rows with errors or unusable keys skipped.'+
    (examples.length?' Example row groups: '+examples.join(' | ')+'.':'')+' Review matches before deciding which records to keep.';
}

function analyseColumn(call,d) {
  const col=datasetColumns(d.values).findIndex(h=>h.toLowerCase()===call.measure.trim().toLowerCase());
  if (col<0) throw Error('Specify an existing column header to analyse.');
  let blank=0,numeric=0,text=0,other=0,errors=0,min=Infinity,max=-Infinity,mean=0;
  for (let i=1;i<d.values.length;i++) {
    const row=d.values[i],v=row[col];
    if (!row.some(x=>x!==''&&x!==null)) continue;
    if (d.errors[i]?.includes(col)) errors++;
    else if (v===''||v===null) blank++;
    else if (typeof v==='number'&&Number.isFinite(v)) {
      numeric++;min=Math.min(v,min);max=Math.max(v,max);mean=mean*((numeric-1)/numeric)+v/numeric;
    } else if (typeof v==='string') text++;
    else other++;
  }
  return call.measure+' in '+source(d)+': '+numeric+' numbers, '+text+' text, '+blank+' blank, '+errors+' errors, '+other+' other.'+
    (numeric?' Numeric range: '+min+' to '+max+'. Mean: '+mean.toLocaleString(undefined,{maximumFractionDigits:10})+'.':'');
}

export async function executeTool(call,context) {
  const target=toolTarget(call,context.sheets,context.selection);
  switch (call.tool) {
    case 'AggregateData': case 'FindDuplicates': case 'AnalyseColumn': {
      const dataset=await context.readDataset(target?.address?target:null);
      const local=call.tool==='AggregateData'?groupedReply(call,dataset):
        call.tool==='FindDuplicates'?duplicates(call,dataset):analyseColumn(call,dataset);
      return {local};
    }
    case 'CompareDatasets': return {compare:true};
    case 'GenerateChartProposal': return {proposal:'chart'};
    case 'GenerateFormattingProposal': return {proposal:'style'};
    case 'GenerateDashboardProposal': return {proposal:'dashboard'};
    case 'InspectWorkbook':
      return {evidence:{worksheets:context.sheets.slice(0,30).map(s=>({name:shorten(s.name),visibility:s.visibility})),
        totalWorksheets:context.sheets.length, sampled:context.snapshots.slice(0,12).map(s=>({sheet:shorten(s.name),
          rows:s.values.length, columns:s.values[0].length, partial:s.partial, address:a1(s.row,s.column,s.values.length,s.values[0].length)}))}};
    default: {
      const chosen=target?.address?await context.readRange(target):
        call.tool==='InspectWorksheet'?context.snapshots.find(s=>s.id===(target?.sheetId??context.selection?.id)):context.selection?.values?context.selection:null;
      if (!chosen) throw Error('Select a range or run a workbook review first.');
      const address=a1(chosen.row,chosen.column,chosen.values.length,chosen.values[0].length);
      const evidence={sheet:shorten(chosen.name),address,rows:chosen.values.length,columns:chosen.values[0].length,
        scope:chosen.partial?'bounded worksheet sample':'range only'};
      if (['ExplainFormula','AnalyseFormulas'].includes(call.tool)) {
        evidence.formulas=[];
        for(let r=0;r<chosen.values.length;r++) for(let c=0;c<chosen.values[r].length;c++)
          if(evidence.formulas.length<12 && formula(chosen,r,c)) {
            const expression=chosen.formulas[r][c];
            evidence.formulas.push({cell:a1(chosen.row+r,chosen.column+c),expression:shorten(expression),
              expressionTruncated:expression.length>180,current:shorten(chosen.display[r][c]),type:chosen.types[r][c],
              references:formulaReferences(expression.slice(0,1024)).slice(0,5)});
          }
      } else {
        evidence.columns=profile(chosen).columns.slice(0,8).map(c=>({...c,label:shorten(c.label)}));
        if (call.tool==='InspectRange') evidence.sample=chosen.values.slice(0,6).map(row=>row.slice(0,6).map(shorten));
      }
      return {evidence};
    }
  }
}
