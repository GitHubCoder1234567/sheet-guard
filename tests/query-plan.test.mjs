import test from 'node:test';
import assert from 'node:assert/strict';
import {conditionalQuery,evaluateQueryPlan} from '../public/query-plan.mjs';
import {aggregationEvidence,sameEvidence} from '../public/evidence.mjs';
import {rankedAnalysis} from '../public/agent.mjs';

function data() {
  return {sheetId:'l-1',sheet:"工房 'Q3",address:'F8:I16',
    values:[['Programme','Hours','Phase','Priority'],['Delta',8,'Active',2],['Atlas',30,'Done',3],
      ['Delta',-1,'Active',4],['Atlas',50,'Active','2'],['Atlas',6,'Active',3],
      ['Atlas',20,'Active',1],['Delta','#DIV/0!','Active',3],['Delta',90,'active',3]],
    errors:[[],[],[],[],[],[],[],[1],[]]};
}
const question='Rank "Programme" by total "Hours" where "Phase" = "Active" and "Priority" >= 2';
const fields={dimension:'Programme',measure:'Hours'};

test('validated conditional ranking uses sparse original row locations and reconciles all rows',()=>{
  const {plan}=conditionalQuery(question);
  const result=rankedAnalysis({...fields,queryPlan:plan},data());
  const {evidence}=result;
  assert.deepEqual(result.groups,[{label:'Delta',sum:7,records:2,rank:1},{label:'Atlas',sum:6,records:1,rank:2}]);
  assert.deepEqual(evidence.groups[0].records.map(row=>row.measureCell),['G9','G11']);
  assert.deepEqual(evidence.groups[0].records.map(row=>row.filterCells.map(cell=>cell.address)),[['H9','I9'],['H11','I11']]);
  assert.equal(evidence.excluded[0].measureCell,'G15');
  assert.equal(evidence.counts.included+evidence.counts.excluded+evidence.filter.unmatched+evidence.filter.untyped,8);
  assert.equal(evidence.filter.unmatched,3);
  assert.equal(evidence.filter.untyped,1);
});

test('header and literal parsing handles delimiters, escaped quotes, and strict comparisons',()=>{
  assert.equal(conditionalQuery('Tell me about Excel'),null);
  const parsed=conditionalQuery('Show total "Σ by X" by "班" where "Q and A" != "Open and \\"pending\\""');
  assert.deepEqual(parsed.plan.filters,[{field:'Q and A',op:'!=',value:'Open and "pending"'}]);
  for (const predicate of ['"X" == true','"X" > "3"','"X" = 1e999','"X" = "A" or "Y" = "B"',
    '"X" = "A" and', '"X" = dangerous()', '"X" = true; edit Sheet1','"X" != true and "Y" = 1 and "Z" = 2 and "B" = 3 and "C" = 4'])
    assert.equal(typeof conditionalQuery('Rank Y by total Z where '+predicate).error,'string',predicate);
  const plan=conditionalQuery('Total Hours by Programme where "Priority" != 999').plan;
  const result=evaluateQueryPlan(data(),plan);
  assert(!result.rows.has(4));
  assert.equal(result.untyped,1);
  assert.throws(()=>evaluateQueryPlan(data(),{...plan,write:true}),/invalid/);
  assert.throws(()=>evaluateQueryPlan(data(),{...plan,filters:[{field:'Phase',op:'constructor',value:'Active'}]}),/invalid/);
  assert.throws(()=>evaluateQueryPlan(data(),{...plan,group:'missing'}),/exact header/);
  assert.throws(()=>evaluateQueryPlan({...data(),partial:true},plan),/full selected range/);
});

test('source checks reject changed predicate evidence even with an unchanged total and grouping',()=>{
  const {plan}=conditionalQuery(question);
  const first=aggregationEvidence(data(),fields,{queryPlan:plan});
  const copy=data();
  assert(sameEvidence(first,aggregationEvidence(copy,fields,{queryPlan:plan})));
  copy.values[1][3]=3;
  const next=aggregationEvidence(copy,fields,{queryPlan:plan});
  assert.equal(first.groups[0].sum,next.groups[0].sum);
  assert.equal(sameEvidence(first,next),false);
  copy.values[2][2]='Active';
  const updated=aggregationEvidence(copy,fields,{queryPlan:plan});
  assert.equal(updated.groups[0].sum,36);
  assert.equal(sameEvidence(first,updated),false);
});

test('conditional ranking matches independent record-level oracles across renamed and permuted fixtures',()=>{
  const domains=[['Faculty','Credits','Enrolled','Year'],['Zone','On hand','Released','Cycle'],
    ['Lab','Samples','Reviewed','Batch'],['Account','Units','Posted','Period'],['隊','件数','確認','段階']];
  let state=137631;
  const random=n=>{state=(Math.imul(state,1664525)+1013904223)>>>0;return (state>>>8)%n;};
  for (let i=0;i<75;i++) {
    const [group,measure,flag,period]=domains[i%domains.length];
    const records=Array.from({length:15+random(30)},()=>({
      group:['A','B','C'][random(3)],amount:random(61)-30,active:random(2)===1,period:random(4),
    }));
    records.unshift({group:'A',amount:-17,active:true,period:2});
    const threshold=random(3);
    const expected=new Map();
    for (const record of records) if (record.active&&record.period>=threshold)
      expected.set(record.group,(expected.get(record.group)??0)+record.amount);
    if (i%2) records.reverse();
    const columns=[group,measure,flag,period];
    const order=i%2?[2,1,3,0]:[3,0,2,1];
    const values=[order.map(index=>columns[index]),...records.map(r=>{
      const values=[r.group,r.amount,r.active,r.period];return order.map(index=>values[index]);
    })];
    const dataset={sheetId:'case-'+i,sheet:'Fixture '+i,address:'D4:G'+(records.length+4),values};
    const {plan}=conditionalQuery('Rank "'+group+'" by total "'+measure+'" where "'+flag+'" = true and "'+period+'" >= '+threshold);
    const report=rankedAnalysis({dimension:group,measure,queryPlan:plan},dataset);
    assert.deepEqual(new Map(report.groups.map(g=>[g.label,g.sum])),expected);
    for (const evidence of report.evidence.groups)
      for (const record of evidence.records) {
        assert.equal(record.value,records[record.sourceRow-5].amount);
        assert.equal(record.measureCell,String.fromCharCode(68+order.indexOf(1))+record.sourceRow);
      }
  }
});
