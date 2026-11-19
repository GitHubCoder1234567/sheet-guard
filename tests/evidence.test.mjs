import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {aggregationEvidence,evidenceGroup,evidenceIntent,sameEvidence} from '../public/evidence.mjs';
import {rankedAnalysis,localQuestion} from '../public/agent.mjs';
import {audit} from '../public/engine.mjs';

const fixture=JSON.parse(await readFile(new URL('./fixtures/sales-audit.json',import.meta.url)));
const fields={dimension:'Category',measure:'Sales'};
const build=(data=fixture,normalised=false)=>aggregationEvidence(data,fields,{sessionId:'session',normalised});

test('Sales Audit exact totals reconcile with the complete row ledger, types, references, and scope',()=>{
  const e=build();
  const expected={Electronics:4780,electronics:2125,Computing:4006,Appliances:3296,Audio:2687,Accessories:421};
  assert.deepEqual(Object.fromEntries(e.groups.map(x=>[x.label,x.sum])),expected);
  assert.equal(e.source.groupRange,'B2:B17');assert.equal(e.source.measureRange,'E2:E17');
  assert.equal(e.counts.repeatedIdRows,1);assert.equal(e.reasons['Blank measure'],1);
  assert.deepEqual(e.excluded.map(x=>x.measureCell),['E13']);
  const contributors=evidenceGroup(e,'Which records contributed to Electronics = 4780?');
  assert.deepEqual(contributors.records.map(r=>[r.groupCell,r.measureCell,r.identifier,r.value]),[
    ['B2','E2','SG-1001',3400],['B3','E3','SG-1002',1380],
  ]);
  assert.equal(e.groups.find(x=>x.label==='Audio').records.find(x=>x.measureCell==='E8').format,'AUD 0');
  assert.equal(fixture.formats[7][3],'USD 0');
  assert.deepEqual(fixture.errors[11],[7]);
  assert.equal(fixture.values[12][6],'31/02/2026');
  assert.deepEqual(fixture.values[15],fixture.values[16]);
});

test('normalised grouping is opt-in; aggregate follow-ups resolve its records, not a 4780-valued cell',()=>{
  const exact=build(), normalized=build(fixture,true);
  assert.equal(evidenceIntent('But there is also lowercase electronics.'),'variants');
  assert.equal(evidenceIntent('Where exactly did 4,780 come from?'),'records');
  assert.equal(evidenceIntent('Show me the rows behind the total.'),'records');
  assert.equal(evidenceGroup(exact,'Where exactly did 4,780 come from?').label,'Electronics');
  assert.deepEqual(evidenceGroup(exact,'Show me the rows behind the total.','Electronics').records.map(r=>r.measureCell),['E2','E3']);
  const group=normalized.groups.find(g=>g.label==='electronics');
  assert.equal(group.sum,6905);assert.deepEqual(group.records.map(r=>r.measureCell),['E2','E3','E4']);
  assert.equal(exact.variants.get('electronics').get('electronics')[0].groupCell,'B4');
  const intent=localQuestion('Which product category has the highest sales? Show the exact calculations and source rows.');
  assert.equal(rankedAnalysis(intent,fixture).groups[0].sum,4780);
});

test('a changed contributing value, label, or order ID invalidates prior evidence',()=>{
  const prior=build();
  for (const [r,c,value] of [[1,4,3500],[2,1,'Audio'],[1,0,'SG-9999']]) {
    const revised=structuredClone(fixture);
    revised.values[r][c]=value;
    assert(!sameEvidence(prior,build(revised)));
  }
  assert(sameEvidence(prior,build()));
});

test('fixture audit distinguishes repeated identifiers, exact duplicates and an unrelated formula error',()=>{
  const types=fixture.values.map((row,i)=>row.map((v,j)=>fixture.errors[i].includes(j)?'Error':
    typeof v==='number'?'Double':v===''?'Empty':'String'));
  const findings=audit([{id:fixture.sheetId,name:fixture.sheet,row:0,column:0,values:fixture.values,types,
    formats:fixture.formats,display:fixture.display,formulas:fixture.values,r1c1:fixture.values,chartCount:0}]);
  assert(findings.some(f=>f.title==='Duplicate records'));
  assert(findings.some(f=>f.title==='Repeated identifier · Order ID'));
  assert(findings.some(f=>f.title==='Excel error values'&&f.detail.includes('H12')));
});

test('format observations block explicitly mixed measure currencies, not unrelated price formats',()=>{
  const intent=localQuestion('Which product category has the highest sales?');
  assert.equal(rankedAnalysis(intent,fixture).groups[0].sum,4780);
  const mixed=structuredClone(fixture);mixed.formats[7][4]='USD 0';
  assert.throws(()=>rankedAnalysis(intent,mixed),/more than one currency/);
});

test('imperative named-sheet ranking resolves the user’s exact query before cloud tool selection',()=>{
  const query='Analyse the Sales Audit worksheet. Rank every product category by total sales. Show the exact cells used for each calculation.';
  const intent=localQuestion(query);
  assert.equal(intent.sheetName,'Sales Audit');
  assert.equal(intent.explainVerified,true);
  const result=rankedAnalysis(intent,fixture);
  assert.equal(result.groups[0].sum,4780);
  assert.deepEqual(result.evidence.groups[0].records.map(r=>r.measureCell),['E2','E3']);
});
