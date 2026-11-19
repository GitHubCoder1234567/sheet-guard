import {a1,rect,DATASET_CELL_LIMIT} from './engine.mjs';

export function rangeShape(dataset) {
  const bounds=rect(dataset.address,DATASET_CELL_LIMIT);
  if (!Array.isArray(dataset.values)||dataset.values.length!==bounds.rows||
      dataset.values.some(row=>!Array.isArray(row)||row.length!==bounds.columns) ||
      dataset.errors && (dataset.errors.length!==bounds.rows || dataset.errors.some(row=>!Array.isArray(row) ||
        row.some(i=>!Number.isInteger(i)||i<0||i>=bounds.columns))))
    throw Error('The workbook snapshot does not match the declared source range.');
  return bounds;
}

export function cellRecord(dataset,bounds,row,column) {
  const raw=dataset.values[row][column];
  const reported=dataset.types?.[row]?.[column]??null;
  const isError=dataset.errors?.[row]?.includes(column)||reported==='Error';
  const kind=isError?'error':reported==='Empty'||raw===null?'blank':
    raw===''?(reported==='String'?'empty_text':'blank'):
    typeof raw==='number'&&Number.isFinite(raw)?'number':
    typeof raw==='string'?'text':typeof raw==='boolean'?'boolean':'unsupported';
  return {address:a1(bounds.r+row,bounds.c+column),rowNumber:bounds.r+row+1,raw,
    displayed:dataset.display?.[row]?.[column]??null,format:dataset.formats?.[row]?.[column]??null,
    formula:dataset.formulas?.[row]?.[column]??null,kind,excelType:reported};
}

export function explicitCurrencyCodes(formats) {
  const currency=/\b(?:AUD|USD|CAD|GBP|EUR|JPY|NZD|CHF|INR|CNY)\b/g;
  return [...new Set(formats.flatMap(format=>format?.match(currency)??[]))];
}
