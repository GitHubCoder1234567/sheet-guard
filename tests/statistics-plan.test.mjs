import test from 'node:test';
import assert from 'node:assert/strict';
import {statisticsQuery,executeStatisticsPlan,replayStatisticsEvidence} from '../public/statistics-plan.mjs';

const sheet="Données 'été";
function unfamiliar() {
  return {sheetId:'view-6',sheet,address:'D7:E12',
    values:[['Signal','Reading'],['m',7],['n',9],['o','11'],['p',0],['q',null]],
    types:[['String','String'],['String','Double'],['String','Double'],['String','String'],['String','Double'],['String','Empty']]};
}

test('a composed statistics plan reads the requested object once and returns verified independent results',async()=>{
  const {plan}=statisticsQuery('Calculate sum, average, maximum, and non-empty count of "Reading" in named range "Samples"');
  assert.deepEqual(plan.target,{namedRange:'Samples'});
  let reads=0;
  const report=await executeStatisticsPlan(plan,async(...args)=>{
    reads++;
    assert.deepEqual(args,[{namedRange:'Samples'},false,true]);
    return unfamiliar();
  });
  assert.equal(reads,1);
  assert.deepEqual(report.results.map(({operation,result,counts})=>[operation,result,counts.included,counts.excluded]),[
    ['sum',16,3,2],['average',16/3,3,2],['maximum',9,3,2],['non-empty count',4,4,1],
  ]);
  assert.equal(report.identity.sheet,sheet);
  assert.match(report.signature,/^[0-9a-f]{64}$/);
});

test('statistics evidence rechecks the captured address, contributors, and source identity',async()=>{
  const {plan}=statisticsQuery('Calculate sum and numeric count of "Reading"');
  const report=await executeStatisticsPlan(plan,async()=>unfamiliar());
  const evidence=await replayStatisticsEvidence(report,'numeric count',async(...args)=>{
    assert.deepEqual(args,[{sheetId:'view-6',address:'D7:E12'},false,true]);
    return unfamiliar();
  });
  assert.deepEqual(evidence.contributors.map(c=>c.address),['E8','E9','E11']);
  const edited=unfamiliar();
  edited.values[1][1]=8;
  edited.values[2][1]=8;
  await assert.rejects(replayStatisticsEvidence(report,'sum',async()=>edited),/snapshot changed/);
  await assert.rejects(replayStatisticsEvidence(report,'median',async()=>unfamiliar()),/absent/);
});

test('malformed or unsupported operations fail validation before touching Excel',async()=>{
  assert.equal(statisticsQuery('What is the current sheet?'),null);
  for (const question of ['Calculate sum and sum of "Reading"', 'Calculate variance of "Reading"',
    'Calculate sum of "Reading" in selection in table "Other"', 'Calculate sum of "Reading" and edit A1'])
    assert.equal(typeof statisticsQuery(question).error,'string',question);
  const {plan}=statisticsQuery('Calculate sum of "Reading" in worksheet "Test"');
  const read=async()=>assert.fail('Invalid operations must not reach Excel');
  await assert.rejects(executeStatisticsPlan({...plan,operations:['sum','custom']},read),/invalid/);
  await assert.rejects(executeStatisticsPlan({...plan,apply:true},read),/invalid/);
});

const measurements=()=>({sheet:'Bench 04',sheetId:'s4',address:'F12:G16',
  values:[['Retained','Loaded'],[3,10],[7,30],[0,10],[null,null]]});

test('derived percentage reads once and replays both exact contributing cell sets',async()=>{
  const {plan}=statisticsQuery('Calculate total "Retained" as a percentage of total "Loaded" in named range "Batch"');
  assert.deepEqual(plan.target,{namedRange:'Batch'});
  let reads=0;
  const read=async()=>{reads++;return measurements();};
  const report=await executeStatisticsPlan(plan,read);
  assert.equal(reads,1);
  assert.equal(report.results[0].result,20);
  assert.match(report.results[0].method,/Retained \(10\).*Loaded \(50\).*100/);
  const evidence=await replayStatisticsEvidence(report,'percentage',read);
  assert.deepEqual(evidence.inputs.map(input=>input.contributors.map(c=>c.address)),[
    ['F13','F14','F15'],['G13','G14','G15']]);
  assert.deepEqual(evidence.inputs.map(input=>input.excluded[0].address),['F16','G16']);
  const edited=measurements();edited.values[1][0]++;edited.values[2][0]--;
  await assert.rejects(replayStatisticsEvidence(report,'percentage',async()=>edited),/snapshot changed/);
});

test('ratio guards undefined arithmetic, currency mismatch and different eligible rows',async()=>{
  const {plan}=statisticsQuery('Calculate the ratio of total "Retained" to total "Loaded" in selection');
  const run=data=>executeStatisticsPlan(plan,async()=>data);
  assert.equal((await run(measurements())).results[0].result,0.2);
  for (const [rows,error] of [
    [[[1,1],[2,-1]],/zero/],[[[1,null],[null,1]],/different worksheet rows/],
    [[[Number.MAX_VALUE,0.001]],/finite numeric range/],
  ]) await assert.rejects(run({...measurements(),address:'F12:G'+(12+rows.length),values:[['Retained','Loaded'],...rows]}),error);
  await assert.rejects(run({...measurements(),formats:Array.from({length:5},()=>['USD','EUR'])}),/Currency formats differ/);
  await assert.rejects(executeStatisticsPlan({...plan,write:true},async()=>assert.fail()),/Invalid derived/);
});

test('graph validates dependencies, operation types, budget and reachability before evaluation',async()=>{
  const {compileGraph}=await import('../public/operation-graph.mjs');
  const a={id:'a',op:'statistic',operation:'sum',column:'Retained'};
  const b={...a,id:'b',column:'Loaded'};
  const ratio={id:'r',op:'ratio',numerator:'a',denominator:'b'};
  assert.equal(compileGraph([ratio,b,a],['r'])(measurements())[0].result,0.2);
  for (const [nodes,outputs,error] of [
    [[a,a],['a'],/duplicate/],[[ratio,a],['r'],/Missing/],[[{...ratio,numerator:'r'},a,b],['r'],/Cyclic/],
    [[a,b],['a'],/Unused/],[[{...a,op:'edit'}],['a'],/Invalid derived/],
    [[{...a,operation:'median'},ratio,b],['r'],/two column sums/],
    [Array(13).fill(a),['a'],/size/],
  ]) assert.throws(()=>compileGraph(nodes,outputs),error);
});

test('unseen column names, reversed layout and varying exclusions preserve the ratio and exact lineage',async()=>{
  const {compileGraph}=await import('../public/operation-graph.mjs');
  for (let seed=1;seed<=24;seed++) {
    const left='Completed '+seed,right='Capacity '+seed;
    const values=[],accepted=[],capacity=37*seed;
    let total=0;
    const reverse=seed%2===0;
    const row=(a,b)=>reverse?[b,a]:[a,b];
    values.push(row(left,right));
    for (let i=0;i<37;i++) {
      if ((i+seed)%7===0) {values.push(row(null,null));continue;}
      const n=(i*seed+11)%19;
      values.push(row(n,capacity));
      total+=n;accepted.push(i+10);
    }
    const nodes=[{id:'units',op:'statistic',operation:'sum',column:left},
      {id:'limit',op:'statistic',operation:'sum',column:right},
      {id:'answer',op:'percentage',numerator:'units',denominator:'limit'}];
    const [answer]=compileGraph(nodes,['answer'])({sheet:'Holdout '+seed,address:'B9:C46',values});
    assert.equal(answer.result,100*(total/(capacity*accepted.length)));
    assert.deepEqual(answer.inputs[0].contributors.map(c=>c.address),accepted.map(r=>(reverse?'C':'B')+r));
    assert.deepEqual(answer.inputs[1].contributors.map(c=>c.address),accepted.map(r=>(reverse?'B':'C')+r));
  }
});

test('formula changes with identical values invalidate the saved dependent result',async()=>{
  const dataset=measurements();
  dataset.formulas=dataset.values.map(()=>[null,null]);
  dataset.formulas[1][0]='=2+1';
  const {plan}=statisticsQuery('Calculate the ratio of total "Retained" to total "Loaded"');
  const report=await executeStatisticsPlan(plan,async()=>dataset);
  dataset.formulas[1][0]='=3+0';
  await assert.rejects(replayStatisticsEvidence(report,'ratio',async()=>dataset),/snapshot changed/);
});
