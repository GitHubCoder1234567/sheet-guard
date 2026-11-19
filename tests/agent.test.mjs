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

test('exact multi-part worksheet request ranks every category with both column references',async()=>{
  const {localQuestion,rankedAnalysis,explicitAudit}=await import('../public/agent.mjs');
  const question='Analyse the current worksheet and identify which product category generates the highest total sales. Calculate the total sales for every category, rank them, explain the calculations, and identify the source cells.';
  const intent=localQuestion(question);
  assert.equal(intent.kind,'rank');assert.equal(intent.detailed,true);assert.equal(intent.worksheet,true);
  assert.equal(explicitAudit(question),false);
  const values=[['Category','Total Sales (AUD)'],['A',12],['B',30],['A',25],['C',7],['B','#VALUE!']];
  const report=rankedAnalysis(intent,{sheet:'Orders',address:'C4:D9',values,errors:[[],[],[],[],[],[1]]});
  assert.deepEqual(report.groups.map(g=>[g.label,g.sum]),[['A',37],['B',30],['C',7]]);
  assert.match(report.text,/Orders!D5:D9.*Orders!C5:C9/s);
  assert.match(report.text,/4 rows included, 1 rows omitted/);
});

test('intent guard distinguishes instructions, auditing, financial objectives and unsupported statistics',async()=>{
  const {calculationGoal,explicitAudit,calculatesGoal,localQuestion}=await import('../public/agent.mjs');
  assert.equal(calculationGoal('Hello!'),null);
  assert.equal(calculationGoal('How many rows does Excel support?'),null);
  assert.equal(calculationGoal('Calculate median sales in this worksheet'), 'unsupported');
  assert.equal(calculationGoal('Calculate the total sales in this worksheet'),'total');
  assert(!calculatesGoal('total',{tool:'FindDuplicates'}));
  assert(!calculatesGoal('ranking',{tool:'AggregateData',operation:'count'}));
  assert(explicitAudit('Find formula errors'));
  assert(!explicitAudit('Analyse the worksheet and rank sales'));
  for (const [q,kind] of [['What is the current worksheet called?','sheet'],['Count rows','count'],
    ['List column headers','columns'],['Explain this formula','formula'],['Find formula errors','errors']])
    assert.equal(localQuestion(q)?.kind,kind);
});
