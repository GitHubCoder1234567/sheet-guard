import {a1,datasetColumns} from './engine.mjs';
import {cellRecord,explicitCurrencyCodes,rangeShape} from './workbook-data.mjs';

const arithmetic=new Set(['sum','average','minimum','maximum','median']);
const operations=new Set([...arithmetic,'numeric count','non-empty count','distinct count']);

function checkedSum(numbers) {
  let sum=0, compensation=0;
  for (const n of numbers) {
    const corrected=n-compensation, next=sum+corrected;
    compensation=(next-sum)-corrected;sum=next;
    if (!Number.isFinite(sum)) throw Error('The total exceeds Excel’s finite numeric range.');
  }
  return sum;
}

export function calculateColumn(dataset,{operation,column}) {
  if (!operations.has(operation)) throw Error('This column operation is unavailable.');
  const bounds=rangeShape(dataset);
  const headers=datasetColumns(dataset.values);
  const index=headers.findIndex(h=>h.toLowerCase()===column.trim().toLowerCase());
  if (index<0) throw Error('Column “'+column+'” was not found. Check the selected header.');
  const contributors=[],excluded=[],distinct=new Set();
  for (let r=1;r<bounds.rows;r++) {
    const cell=cellRecord(dataset,bounds,r,index);
    let reason;
    if (operation==='non-empty count') {
      if (cell.kind==='blank') reason='blank';
    } else if (operation==='distinct count') {
      if (!['text','empty_text','number','boolean'].includes(cell.kind)) reason=cell.kind;
      else if (distinct.has(cell.raw)) reason='repeated value';
      else distinct.add(cell.raw);
    } else if (cell.kind!=='number') reason=cell.kind;
    if (reason) excluded.push({...cell,reason});else contributors.push(cell);
  }
  if (arithmetic.has(operation)&&!contributors.length) throw Error('No numeric values found in the requested measure column.');
  const codes=explicitCurrencyCodes(contributors.map(x=>x.format));
  if (arithmetic.has(operation)&&codes.length>1)
    throw Error('The measure has more than one currency format. Select a single-currency range or a converted column.');
  let value=contributors.length;
  if (arithmetic.has(operation)) {
    const nums=contributors.map(x=>x.raw);
    if (operation==='sum') value=checkedSum(nums);
    if (operation==='average') value=checkedSum(nums)/nums.length;
    if (operation==='minimum') value=nums.reduce((low,n)=>Math.min(low,n),Infinity);
    if (operation==='maximum') value=nums.reduce((high,n)=>Math.max(high,n),-Infinity);
    if (operation==='median') {
      nums.sort((a,b)=>a-b);const i=Math.floor(nums.length/2);
      value=nums.length%2?nums[i]:nums[i]===nums[i-1]?nums[i]:nums[i]/2+nums[i-1]/2;
    }
  }
  if (!Number.isFinite(value)||contributors.length+excluded.length!==bounds.rows-1)
    throw Error('The calculation could not be verified against the source cells.');
  return {schema:1,operation,field:headers[index],result:value,
    source:{sheet:dataset.sheet,sheetId:dataset.sheetId,table:dataset.table??null,
      range:a1(bounds.r+1,bounds.c+index,bounds.rows-1,1)},
    counts:{included:contributors.length,excluded:excluded.length},contributors,excluded,
    scope:'full selected rectangle; includes hidden and filtered rows',
    observedCurrency:codes.length===1?codes[0]:null};
}
