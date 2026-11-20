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
