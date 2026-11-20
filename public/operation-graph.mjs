import {calculateColumn,supportedColumnOperation} from './calculation.mjs';
import {exactKeys,fieldName} from './query-plan.mjs';

// Validate and order the entire read-only graph before acquiring a workbook snapshot.
export function compileGraph(nodes,outputs) {
  const validId=id=>typeof id==='string' && /^[a-z][a-z0-9_]{0,31}$/.test(id);
  if (!Array.isArray(nodes)||!nodes.length||nodes.length>12||
      !Array.isArray(outputs)||!outputs.length||outputs.length>nodes.length||new Set(outputs).size!==outputs.length)
    throw Error('Invalid calculation graph size or outputs.');
  const byId=new Map();
  for (const node of nodes) {
    if (!validId(node?.id)||byId.has(node.id)) throw Error('Invalid or duplicate calculation ID.');
    if (node.op==='statistic') {
      if (!exactKeys(node,['id','op','column','operation'])||!fieldName(node.column)||!supportedColumnOperation(node.operation))
        throw Error('Invalid statistic operation.');
    } else if (!['ratio','percentage'].includes(node.op)||!exactKeys(node,['id','op','numerator','denominator'])||
        !validId(node.numerator)||!validId(node.denominator)) throw Error('Invalid derived operation.');
    byId.set(node.id,{...node});
  }
  const order=[],visited=new Set(),visiting=new Set();
  function visit(id) {
    if (visited.has(id)) return;
    if (visiting.has(id)) throw Error('Cyclic calculation dependency.');
    const node=byId.get(id);
    if (!node) throw Error('Missing calculation dependency.');
    visiting.add(id);
    if (node.op!=='statistic') {
      visit(node.numerator);
      visit(node.denominator);
      if ([node.numerator,node.denominator].some(input=>byId.get(input).operation!=='sum'))
        throw Error('Ratios currently require two column sums.');
    }
    visiting.delete(id);
    visited.add(id);
    order.push(node);
  }
  outputs.forEach(visit);
  if (visited.size!==nodes.length) throw Error('Unused calculation operation.');
  const requested=[...outputs];
  return dataset=>{
    const values=new Map();
    for (const node of order) {
      if (node.op==='statistic') {
        values.set(node.id,calculateColumn(dataset,node));
        continue;
      }
      const numerator=values.get(node.numerator),denominator=values.get(node.denominator);
      if (numerator.observedCurrency!==denominator.observedCurrency)
        throw Error('Currency formats differ. Use totals in the same units.');
      if (numerator.counts.included!==denominator.counts.included||
          numerator.contributors.some((cell,index)=>cell.rowNumber!==denominator.contributors[index].rowNumber))
        throw Error('The totals include different worksheet rows. Choose two measures with matching numeric rows.');
      if (denominator.result===0) throw Error('The denominator total is zero. The requested ratio is undefined.');
      const quotient=numerator.result/denominator.result;
      const result=node.op==='percentage'?quotient*100:quotient;
      if (!Number.isFinite(result)) throw Error('The ratio exceeds the finite numeric range.');
      values.set(node.id,{operation:node.op,result,field:numerator.field+' / '+denominator.field,
        counts:numerator.counts,observedCurrency:null,inputs:[numerator,denominator],
        method:'Sum of '+numerator.field+' ('+numerator.result+') divided by sum of '+denominator.field+
          ' ('+denominator.result+')'+(node.op==='percentage'?', multiplied by 100.':'.')});
    }
    return requested.map(id=>values.get(id));
  };
}
