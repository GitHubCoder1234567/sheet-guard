import {a1,rect,datasetColumns,DATASET_CELL_LIMIT} from './engine.mjs';
import {explicitCurrencyCodes} from './workbook-data.mjs';

const present=value=>value!==''&&value!==null&&value!==undefined;
const numeric=value=>typeof value==='number'&&Number.isFinite(value);

function populatedSpans(flags) {
  const spans=[];
  let start=-1;
  for (let i=0;i<=flags.length;i++) {
    if (flags[i]&&start<0) start=i;
    if (!flags[i]&&start>=0) {spans.push([start,i]);start=-1;}
  }
  return spans;
}

function classifyColumn(label,values,types,formats,formulas) {
  const nonempty=values.filter((value,i)=>present(value)&&types[i]!=='Error');
  const numbers=nonempty.filter(numeric).length;
  const text=nonempty.filter(v=>typeof v==='string').length;
  const evidence={nonempty:nonempty.length,numeric:numbers,text,distinct:new Set(nonempty).size,
    booleans:nonempty.filter(v=>typeof v==='boolean').length,errors:types.filter(v=>v==='Error').length,
    formulas:formulas.filter((v,i)=>typeof v==='string'&&v.startsWith('=')&&v!==values[i]).length,
    currencyCodes:explicitCurrencyCodes(formats)};
  const key=/\b(?:id|identifier|key)\b/i.test(label);
  if (key) return {role:'identifier candidate',confidence:evidence.distinct===nonempty.length&&nonempty.length>=3?'supported':'tentative',evidence};
  if (/\b(?:date|month|period|year)\b/i.test(label))
    return {role:'date or period candidate',confidence:'tentative',evidence};
  if (numbers===nonempty.length&&numbers>0) return {role:'numeric measure candidate',confidence:'supported',evidence};
  if (text===nonempty.length&&text>0) return {role:'text or dimension candidate',confidence:'supported',evidence};
  return {role:'mixed or empty',confidence:'tentative',evidence};
}

export function discoverRegions({values,address,partial=false,types,formats,formulas}) {
  const bounds=rect(address,DATASET_CELL_LIMIT);
  if (values.length!==bounds.rows||values.some(row=>row.length!==bounds.columns))
    throw Error('Workbook snapshot dimensions could not be verified.');
  const colFlags=new Array(bounds.columns).fill(false);
  for (const row of values) row.forEach((value,c)=>{if(present(value))colFlags[c]=true;});
  const regions=[],ambiguous=[];
  for (const [c,end] of populatedSpans(colFlags)) {
    if (end-c<2) continue;
    const spans=populatedSpans(values.map(row=>row.slice(c,end).some(present)));
    for (const [start,stop] of spans) {
      if (stop-start<2) continue;
      let headers,headerRow;
      for (let r=start;r<Math.min(stop-1,start+6);r++) {
        try {headers=datasetColumns([values[r].slice(c,end)]);headerRow=r;break;}
        catch { /* Titles and incomplete or duplicate headers are not promoted. */ }
      }
      const addressOf=(r,rows)=>a1(bounds.r+r,bounds.c+c,rows,end-c);
      if (headerRow===undefined) {
        ambiguous.push({address:addressOf(start,stop-start),reason:'No distinct complete header within the first six populated rows.'});
        continue;
      }
      let repeated=false,footer=false;
      const following=values.slice(headerRow+1,stop).map(row=>row.slice(c,end));
      const lower=headers.map(h=>h.toLowerCase());
      for (const row of following) {
        if (row.every((v,i)=>typeof v==='string'&&v.trim().toLowerCase()===lower[i])) repeated=true;
        if (typeof row[0]==='string'&&/^(?:grand total|subtotal|total)$/i.test(row[0].trim())) footer=true;
      }
      if (repeated||footer) {
        ambiguous.push({address:addressOf(headerRow,stop-headerRow),reason:repeated?'Repeated header row.':'Possible total or subtotal row.'});
        continue;
      }
      const field=(matrix,i)=>matrix?.slice(headerRow+1,stop).map(row=>row?.[c+i])??[];
      const columns=headers.map((label,i)=>({label,...classifyColumn(label,following.map(row=>row[i]),
        field(types,i),field(formats,i),field(formulas,i))}));
      const numericEvidence=columns.some(x=>x.role==='numeric measure candidate');
      const confidence=!partial&&numericEvidence&&following.length>=2?'supported':'tentative';
      regions.push({address:addressOf(headerRow,stop-headerRow),headers,columns,rows:following.length,confidence,
        offsets:{row:headerRow,column:c,rows:stop-headerRow,columns:end-c}});
    }
  }
  return {regions,ambiguous,partial,coverage:partial?'bounded sample':'full requested rectangle'};
}

export function chooseDiscoveredRegion(dataset) {
  const profile=discoverRegions(dataset);
  if (profile.regions.length!==1 || profile.ambiguous.length || profile.regions[0].confidence!=='supported')
    return null;
  const chosen=profile.regions[0],{row,column,rows,columns}=chosen.offsets;
  const cut=matrix=>matrix?.slice(row,row+rows).map(line=>line.slice(column,column+columns));
  return {...dataset,address:chosen.address,values:cut(dataset.values),display:cut(dataset.display),formats:cut(dataset.formats),types:cut(dataset.types),formulas:cut(dataset.formulas),
    errors:dataset.errors?.slice(row,row+rows).map(line=>line.filter(i=>i>=column&&i<column+columns).map(i=>i-column)),
    discovery:'worksheet region inferred from headers and typed data'};
}
