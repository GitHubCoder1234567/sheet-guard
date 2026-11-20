import {aggregationEvidence} from './evidence.mjs';
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
      const dataset=await context.readDataset(target?.address?target:null,context.detailed&&call.tool==='AggregateData');
      if (call.tool==='AggregateData' && context.detailed && call.measure && call.dimension &&
          ['total','highest','lowest'].includes(call.operation)) {
        const report=rankedAnalysis({...call,operation:context.rankingOperation??call.operation,sessionId:context.sessionId},dataset);
        return {local:report.text,report};
      }
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

// Resolve a concrete analytical objective before generic workbook actions.
export function calculationRequired(question) {
  return /\b(?:calculate|sum|rank|highest|lowest|average|median|how many|grand total|total sales|total revenue)\b/i.test(question) &&
    /\b(?:worksheet|workbook|sales|revenue|expenses|amount|dataset|records|rows|transactions|column|category|categories|product|products|selection)\b/i.test(question) &&
    !/^\s*(?:how (?:do|can|should) i|how many (?:rows|columns) (?:does|can) Excel|show me how to|write (?:an? )?(?:excel )?formula)\b/i.test(question);
}

export function calculationGoal(question) {
  if (!calculationRequired(question)) return null;
  if (/\b(?:median|standard deviation|percentile)\b/i.test(question)) return 'unsupported';
  if (/\b(?:rank|highest|lowest)\b/i.test(question)) return 'ranking';
  if (/\b(?:how many|count)\b/i.test(question)) return 'count';
  if (/\b(?:average|mean)\b/i.test(question)) return 'average';
  return 'total';
}

export function calculatesGoal(goal,plan) {
  if (goal==='ranking') return plan.tool==='AggregateData' && !!plan.measure && !!plan.dimension &&
    ['highest','lowest','total'].includes(plan.operation);
  if (goal==='count') return plan.tool==='AggregateData' && plan.operation==='count';
  if (goal==='average') return plan.tool==='AnalyseColumn' && !!plan.measure;
  if (goal==='total') return plan.tool==='AggregateData' && plan.operation==='total' && !!plan.measure;
  return false;
}

export function explicitAudit(question) {
  return /^\s*(?:please\s+)?(?:audit|review|scan|check)(?:\s+(?:my|this|the|current|active))*\s+(?:workbook|worksheet|sheet|spreadsheet)[.!?]*\s*$/i.test(question) ||
    /\b(?:find|detect|list|report|check for|audit for)\s+(?:(?:the|all|any|potential)\s+)*(?:formula errors|calculation errors|broken formulas)\b/i.test(question);
}

function compoundRank(q) {
  const match=/\bwhich ([^.!?;]+?) (?:has|had|generates?|generated|produces?|earned) (?:the )?(highest|lowest) (?:total )?([^.!?;,]+?)(?=\s+(?:and|then)\s+(?:calculate|rank|list|show|explain|identify|provide)\b|[.!?;,]|$)/i.exec(q);
  if (!match) return null;
  return {kind:'rank',dimension:match[1].trim().replace(/^product category$/i,'Product Category'),
    operation:match[2].toLowerCase(),measure:match[3].replace(/\s+(?:with|using|when grouping|if grouped)\s+case.insensitiv(?:e|ely).*$/i,'').trim(),
    normalised:/\b(?:case.insensitiv(?:e|ely)|ignore case|normalise whitespace|normalize whitespace)\b/i.test(q),
    detailed:/\b(?:every|all|rank|explain|source|cells|ranges|limitations)\b/i.test(q),
    worksheet:/\b(?:current|active|this) worksheet\b/i.test(q) && !/\b(?:selected|selection)\b/i.test(q)};
}

function namedWorksheetRanking(q) {
  const ranking=/\brank\s+(?:every |each |all )?([^.!?;]+?)\s+by\s+(?:total\s+)?([^.!?;]+?)(?=[.!?;]|$)/i.exec(q);
  if (!ranking) return null;
  const named=/\b(?:analy[sz]e|inspect|review)\s+(?:the\s+)?(?:"([^"]+)"|“([^”]+)”|([^.!?;]+?))\s+worksheet\b/i.exec(q);
  if (!named) return null;
  const sheetName=(named[1]??named[2]??named[3]).trim();
  const localScope=/^(?:the )?(?:current|active|this)$/i.test(sheetName);
  return {kind:'rank',dimension:ranking[1].trim(),measure:ranking[2].trim(),operation:'highest',
    worksheet:true,sheetName:localScope?null:sheetName,explainVerified:true,detailed:true,
    normalised:/\b(?:case.insensitiv(?:e|ely)|ignore case|normalise whitespace|normalize whitespace)\b/i.test(q)};
}

export function localQuestion(question) {
  const q=question.trim().replace(/[.?!]+$/,'');
  const compound=namedWorksheetRanking(q)??compoundRank(q);
  if (compound) return compound;
  if (/^(?:what(?: is|'s) (?:the )?(?:(?:current|active) )?(?:worksheet|sheet)(?: called| name)?|name (?:the )?(?:active|current) (?:worksheet|sheet))$/i.test(q)) return {kind:'sheet'};
  if (/^(?:(?:find|show|list|check for|report) (?:the |all |any )?formula errors(?: in (?:this|the current) (?:sheet|worksheet))?)$/i.test(q))
    return {kind:'errors'};
  if (/^(?:how many (?:data )?(?:records|rows)(?: (?:are there|are in|in)(?: (?:this|the|my) (?:dataset|selection|table))?)?|count (?:the )?(?:records|rows)(?: in (?:this|the|my) (?:dataset|selection|table))?)$/i.test(q)) return {kind:'count'};
  if (/^(?:what (?:columns|headers) (?:are (?:in|there in) (?:this|the|my) (?:dataset|selection|table)|are there)|list (?:the )?(?:column (?:names|headers)|columns|headers)(?: in (?:this|the|my) (?:dataset|selection|table))?)$/i.test(q)) return {kind:'columns'};
  const rank=/^which (.+?) (?:has|had|generates?) (?:the )?(highest|lowest) (?:total )?(.+)$/i.exec(q);
  if (rank) return {kind:'rank',dimension:rank[1],operation:rank[2].toLowerCase(),measure:rank[3]};
  if (/^explain (?:this|the selected) (?:formula|calculation)$/i.test(q)) return {kind:'formula'};
  if (/^(?:show me |calculate )?total .+ by .+$/i.test(q)) return {kind:'group'};
  return null;
}

export function datasetAnswer(intent,d) {
  const headers=datasetColumns(d.values);
  if (intent.kind==='columns') return headers.length+' columns in '+source(d)+': '+headers.join(', ')+'.';
  if (intent.kind==='count') return groupedReply({operation:'count'},d);
  const fields=resolveRankFields(intent,headers);
  return groupedReply({operation:intent.operation,...fields},d);
}

function resolveRankFields(intent,headers) {
  const lookup=label=>{
    const normal=v=>v.trim().toLowerCase().replace(/^total\s+/,'').replace(/\s+/g,' ');
    const exact=headers.filter(h=>h.toLowerCase()===label.trim().toLowerCase());
    let matches=exact.length?exact:headers.filter(h=>normal(h)===normal(label));
    if (!matches.length && /^(?:product )?categor(?:y|ies)$/i.test(label.trim()))
      matches=headers.filter(h=>/^(?:product )?categor(?:y|ies)$/i.test(h.trim()));
    if (!matches.length && /^(?:sales|revenue)$/i.test(label.trim()))
      matches=headers.filter(h=>normal(h).replace(/\s*\([A-Z]{3}\)$/i,'')===normal(label));
    if (matches.length!==1) throw Error('Which column do you mean by “'+label+'”? Available columns: '+headers.slice(0,12).join(', ')+'.');
    return matches[0];
  };
  return {dimension:lookup(intent.dimension),measure:lookup(intent.measure)};
}


export function rankedAnalysis(intent,d) {
  const headers=datasetColumns(d.values);
  const fields=resolveRankFields(intent,headers);
  const evidence=aggregationEvidence(d,fields,intent);
  if (!evidence.groups.length) throw Error(evidence.filter?
    'No rows both matched the filters and had an eligible grouping value and numeric measure in '+source(d)+'.':
    'No eligible numeric rows in '+source(d)+'. Check the measure and grouping cells.');
  const {counts,source:scope}=evidence;
  const codes=new Set(evidence.observations.measureFormats.flatMap(format=>
    format.match(/\b(?:AUD|USD|NZD|CAD|EUR|GBP|JPY|CHF|CNY|INR)\b/g)??[]));
  if (codes.size>1) throw Error('The measure column shows more than one currency ('+[...codes].join(', ')+
    '). Select a single-currency range or supply a separate converted measure before ranking.');
  const excluded=Object.entries(evidence.reasons).map(([reason,count])=>reason+': '+count).join(', ');
  const dimensionRange=scope.sheet+'!'+scope.groupRange,measureRange=scope.sheet+'!'+scope.measureRange;
  let rank=0;
  const groups=evidence.groups.map(({label,sum,records},index)=>{
    if (index===0 || sum!==evidence.groups[index-1].sum) rank=index+1;
    return {label,sum,records:records.length,rank};
  });
  const best=intent.operation==='lowest'?groups.at(-1).sum:groups[0].sum;
  const winners=groups.filter(g=>g.sum===best);
  const money=n=>n.toLocaleString(undefined,{maximumFractionDigits:10});
  const title=(intent.operation==='lowest'?'Lowest':'Highest')+' total '+fields.measure;
  const top=winners.slice(0,8).map(g=>g.label).join(', ');
  const text=title+': '+top+' · '+money(best)+(winners.length>1?' ('+winners.length+' tied categories)':'')+'.\n\n'+
    'I summed numeric values in '+measureRange+' for each matching value in '+dimensionRange+
    '. Ranked by total, highest first. '+counts.included+' rows included, '+counts.excluded+
    ' rows omitted'+(excluded?' ('+excluded+')':'')+'. Empty rows are excluded. '+
    (intent.normalised?'Text labels are grouped case-insensitively with trimmed, collapsed whitespace. ':
      'Category labels are grouped exactly as entered. ')+
    'Repeated records are counted separately'+(counts.repeatedIdRows?' ('+counts.repeatedIdRows+' repeated identifier occurrence(s))':'')+
    '. The full requested range includes hidden and filtered rows. Scope: '+source(d)+'.'+
    (evidence.filter?' Filter: '+evidence.filter.condition+'. '+evidence.filter.unmatched+' rows did not match and '+
      evidence.filter.untyped+' lacked compatible filter values. Numeric filters compare underlying Excel values.':'');
  return {text,groups,...fields,evidence};
}
