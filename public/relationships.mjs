import {datasetColumns} from './engine.mjs';
import {cellRecord,rangeShape} from './workbook-data.mjs';

const exampleLimit=3;
const referenceLimit=4;
const keyKinds=new Set(['number','text','boolean']);

function keyIndex(dataset,header) {
  if (dataset.partial) throw Error('Capture the full datasets to check key relationships.');
  if (typeof header!=='string') throw Error('Choose a key column for each capture.');
  const bounds=rangeShape(dataset),headers=datasetColumns(dataset.values);
  const col=headers.indexOf(header);
  if (col<0) throw Error('The chosen key column is missing. Capture the data again.');
  const entries=new Map();
  let missing=0,invalid=0,duplicated=0;
  for (let i=1;i<bounds.rows;i++) {
    const {address,kind,raw}=cellRecord(dataset,bounds,i,col);
    if (kind==='blank'||kind==='empty_text'||typeof raw==='string'&&!raw.trim()) {missing++;continue;}
    if (!keyKinds.has(kind)) {invalid++;continue;}
    const entry=entries.get(raw)??{count:0,addresses:[]};
    entry.count++;
    if (entry.count===2) duplicated++;
    if (entry.addresses.length<referenceLimit) entry.addresses.push(address);
    entries.set(raw,entry);
  }
  return {entries,stats:{source:{sheetId:dataset.sheetId,sheet:dataset.sheet,range:dataset.address,header},
    dataRows:bounds.rows-1,distinct:entries.size,duplicated,missing,invalid}};
}

export function inspectKeyRelationship(first,second,{firstKey,secondKey}) {
  const left=keyIndex(first,firstKey),right=keyIndex(second,secondKey);
  const counts={oneToOne:0,oneToMany:0,manyToOne:0,manyToMany:0,firstOnly:0,secondOnly:0};
  const examples=Object.fromEntries(Object.keys(counts).map(kind=>[kind,[]]));
  let matchingPairs=0,firstUnmatched=0,secondUnmatched=0,firstMatched=0,secondMatched=0;
  function record(kind,value,firstEntry=null,secondEntry=null) {
    counts[kind]++;
    if (examples[kind].length<exampleLimit)
      examples[kind].push({value,first:firstEntry,second:secondEntry});
  }
  for (const [value,firstEntry] of left.entries) {
    const secondEntry=right.entries.get(value);
    if (!secondEntry) {
      firstUnmatched+=firstEntry.count;
      record('firstOnly',value,firstEntry);
      continue;
    }
    const a=firstEntry.count,b=secondEntry.count;
    firstMatched+=a;secondMatched+=b;
    matchingPairs+=a*b;
    const kind=a===1?(b===1?'oneToOne':'oneToMany'):(b===1?'manyToOne':'manyToMany');
    record(kind,value,firstEntry,secondEntry);
  }
  for (const [value,entry] of right.entries) {
    if (left.entries.has(value)) continue;
    secondUnmatched+=entry.count;
    record('secondOnly',value,null,entry);
  }
  if (firstMatched+firstUnmatched+left.stats.missing+left.stats.invalid!==left.stats.dataRows||
      secondMatched+secondUnmatched+right.stats.missing+right.stats.invalid!==right.stats.dataRows)
    throw Error('The relationship counts did not reconcile with the two captured ranges.');
  return {schema:1,first:{...left.stats,matched:firstMatched,unmatched:firstUnmatched},
    second:{...right.stats,matched:secondMatched,unmatched:secondUnmatched},counts,matchingPairs,examples};
}
