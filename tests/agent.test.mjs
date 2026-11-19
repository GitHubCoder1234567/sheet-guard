import test from 'node:test';
import assert from 'node:assert/strict';
import {executeTool,toolTarget} from '../public/agent.mjs';

const call=(tool,extra={})=>({tool,sheet:'',address:'',measure:'',dimension:'',operation:'',...extra});
function context(values) {
  return {sheets:[{id:'s',name:'Sales',visibility:'Visible'},{id:'secret',name:'Hidden',visibility:'Hidden'}],
    snapshots:[],selection:{id:'s'},readDataset:async()=>({sheet:'Sales',address:'A1:C'+values.length,
      values,errors:values.map(()=>[])})};
}
test('AI extrema use all selected rows and provide address and coverage',async()=>{
  const data=[['Product','Sales','ID'],...Array.from({length:30000},(_,i)=>[i<5?'Early':'Later',2,i])];
  const result=await executeTool(call('AggregateData',{measure:'Sales',dimension:'Product',operation:'highest'}),context(data));
  assert.match(result.local,/Later · 59[,.]?990/);
  assert.match(result.local,/Sales!A1:C30001/);
  assert.match(result.local,/30000 rows included/);
});
test('record count skips empty rows and duplicate detection reports actual row positions',async()=>{
  const ctx=context([['Product','Sales','ID'],['Item',5,1],['Item',5,1],['','',''],['Other',3,2]]);
  const count=await executeTool(call('AggregateData',{operation:'count'}),ctx);
  assert.match(count.local,/^3 records/);
  const match=await executeTool(call('FindDuplicates',{measure:'ID'}),ctx);
  assert.match(match.local,/1 groups, 1 additional/);
  assert.match(match.local,/2, 3/);
});
test('unsupported tools and invalid or hidden targets are rejected',()=>{
  const ctx=context([['Product','Price']]);
  for (const invalid of [call('ExecuteCode'),call('InspectRange',{sheet:'Hidden',address:'A1'}),
    call('InspectRange',{sheet:'Sales',address:'Book!A1'}),call('InspectRange',{address:'A1:Z10000'})])
    assert.throws(()=>toolTarget(invalid,ctx.sheets,ctx.selection));
});
test('formula evidence ties actual formulas and references to their source cells',async()=>{
  const ctx=context([]);
  ctx.readRange=async()=>({id:'s',name:'Sales',row:2,column:2,values:[[12]],formulas:[['=A3*B3']],
    display:[['12']],types:[['Double']]});
  const result=await executeTool(call('ExplainFormula',{sheet:'Sales',address:'C3'}),ctx);
  assert.equal(result.evidence.formulas[0].cell,'C3');
  assert.deepEqual(result.evidence.formulas[0].references.map(r=>r.first),['A3','B3']);
});
