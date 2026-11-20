import test from 'node:test';
import assert from 'node:assert/strict';
import {inspectKeyRelationship} from '../public/relationships.mjs';

const first={sheetId:'f',sheet:'Requests',address:'B4:C13',
  values:[['Request','Title'],[0,'zero'],[1,'number'],['01','other'],['A','a1'],['A','a2'],
    ['B','b1'],['B','b2'],['C','c1'],['','none']]};
const second={sheetId:'s',sheet:"Réponses '03",address:'G7:H17',
  values:[['Activity','Ticket'],['zero',0],['text one','1'],['a1','A'],['a2','A'],['a3','A'],['b','B'],
    ['c1','C'],['c2','C'],['error','#N/A'],['empty','  ']],
  errors:[[],[],[],[],[],[],[],[],[],[1],[]]};
const keys={firstKey:'Request',secondKey:'Ticket'};

test('relationship checks reconcile typed key records and preserve addresses across two captures',()=>{
  const r=inspectKeyRelationship(first,second,keys);
  assert.deepEqual(r.counts,{oneToOne:1,oneToMany:1,manyToOne:1,manyToMany:1,firstOnly:2,secondOnly:1});
  assert.equal(r.matchingPairs,11);
  assert.equal(r.first.unmatched,2);
  assert.equal(r.first.missing,1);
  assert.equal(r.first.duplicated,2);
  assert.equal(r.second.unmatched,1);
  assert.equal(r.second.missing,1);
  assert.equal(r.second.invalid,1);
  assert.deepEqual(r.examples.manyToMany[0],{value:'A',
    first:{addresses:['B8','B9'],count:2},second:{addresses:['H10','H11','H12'],count:3}});
  assert.deepEqual(r.examples.firstOnly.map(x=>x.value),[1,'01']);
});

test('relationship checks decline incomplete data and a missing key column',()=>{
  assert.throws(()=>inspectKeyRelationship({...first,partial:true},second,keys),/full datasets/);
  assert.throws(()=>inspectKeyRelationship(first,second,{...keys,firstKey:'Unknown'}),/missing/);
  assert.throws(()=>inspectKeyRelationship(first,{...second,values:[]},keys),/declared source/);
});

test('high-multiplicity keys are counted without materializing the Cartesian product',()=>{
  const dataset=(label,rows)=>({sheet:label,sheetId:label,address:'D1:D'+(rows+1),
    values:[[label],...Array.from({length:rows},()=>['same'])]});
  const left=dataset('Reference',20000),right=dataset('Serial',12000);
  const report=inspectKeyRelationship(left,right,{firstKey:'Reference',secondKey:'Serial'});
  assert.equal(report.matchingPairs,240000000);
  assert.equal(report.counts.manyToMany,1);
  assert.equal(report.examples.manyToMany[0].first.addresses.length,4);
  assert.equal(report.examples.manyToMany[0].second.count,12000);
});
