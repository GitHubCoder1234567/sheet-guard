import test from 'node:test';
import assert from 'node:assert/strict';
import {discoverRegions,chooseDiscoveredRegion} from '../public/discovery.mjs';
import {calculateColumn} from '../public/calculation.mjs';
import {aggregationEvidence} from '../public/evidence.mjs';
import {WorkbookError,readDataset,workbookFailure} from '../public/excel.mjs';
import {runLightDiagnostics,runDiagnostics} from '../public/diagnostics.mjs';
import {scenarios,verifyLiveReports} from '../tools/verify-live-report.mjs';
import {corpus,holdout,workbookCase} from './fixtures/workbook-corpus.mjs';

test('discovery accepts a bounded header under a title and refuses separate blocks, repeated headers and totals',()=>{
  const data={address:'C6:D11',values:[['Report',''],['',''],['Group','Sales'],['East',7],['West',9],['East',10]]};
  assert.equal(chooseDiscoveredRegion(data).address,'C8:D11');
  const split={address:'A1:E3',values:[['A','B','','C','D'],['x',3,'','p',4],['y',5,'','q',6]]};
  assert.equal(discoverRegions(split).regions.length,2);
  assert.equal(chooseDiscoveredRegion(split),null);
  for (const last of [['Group','Sales'],['Grand total',17]]) {
    const ambiguous={address:'A1:B4',values:[['Group','Sales'],['east',7],['west',10],last]};
    assert.equal(chooseDiscoveredRegion(ambiguous),null);
    assert.equal(discoverRegions(ambiguous).ambiguous.length,1);
  }
  assert.equal(chooseDiscoveredRegion({...data,partial:true}),null);
});

test('calculation contract distinguishes numeric text, formula empty text, actual blanks, zeros and errors',()=>{
  const data={sheet:'Different sheet',sheetId:'s',address:'D2:D9',
    values:[['Actual'],[0],[6],['2'],[''],[''],['#REF!'],[8]],
    types:[['String'],['Double'],['Double'],['String'],['Empty'],['String'],['Error'],['Double']]};
  const run=operation=>calculateColumn(data,{operation,column:'Actual'});
  assert.equal(run('sum').result,14);
  assert.equal(run('average').result,14/3);
  assert.equal(run('median').result,6);
  assert.equal(run('minimum').result,0);
  assert.equal(run('maximum').result,8);
  assert.equal(run('numeric count').result,3);
  assert.equal(run('non-empty count').result,6);
  assert.equal(run('distinct count').result,5);
  assert.deepEqual(run('sum').contributors.map(c=>c.address),['D3','D4','D9']);
  assert.equal(run('sum').source.range,'D3:D9');
  const formats=Array.from({length:8},()=>['AUD 0']);formats[7]=['USD 0'];
  assert.throws(()=>calculateColumn({...data,formats},{operation:'sum',column:'Actual'}),/more than one currency/);
});

test('corpus and holdout match independent integer oracles and label unsupported discovery correctly',t=>{
  let verified=0, unsupported=0;
  for (const f of [...corpus(),...holdout()]) {
    const d=chooseDiscoveredRegion(f.dataset);
    if (['ambiguous','multiple'].includes(f.expectation)) {
      assert.equal(d,null,f.dataset.sheet);unsupported++;continue;
    }
    assert(d,f.dataset.sheet);
    if (f.expectation==='mixed_currency') {
      assert.throws(()=>calculateColumn(d,{operation:'sum',column:f.headers[2]}),/currency/);
      assert.throws(()=>aggregationEvidence(d,{dimension:f.headers[1],measure:f.headers[2]}),/currency/);
      unsupported++;continue;
    }
    const got=aggregationEvidence(d,{dimension:f.headers[1],measure:f.headers[2]});
    assert.deepEqual(Object.fromEntries(got.groups.map(g=>[g.label,g.sum])),{...f.expected});
    for (const group of got.groups) assert.equal(group.records.reduce((n,r)=>n+r.value,0),group.sum);
    verified++;
  }
  assert.equal(verified+unsupported,50);
  t.diagnostic(verified+' supported fixtures matched the oracle; '+unsupported+' were declined as expected');
});

test('seeded property and metamorphic tests preserve exact sums on reordered source data',()=>{
  for (let i=0;i<90;i++) {
    const f=workbookCase(30391+i*79,i%10,8+i%40);
    const [header,...body]=f.dataset.values;
    const d={...f.dataset,values:[header,...body.reverse()],sheet:'Renamed '+i};
    const result=aggregationEvidence(d,{dimension:f.headers[1],measure:f.headers[2]});
    assert.deepEqual(Object.fromEntries(result.groups.map(g=>[g.label,g.sum])),{...f.expected});
    assert.equal(result.groups.reduce((n,g)=>n+g.records.length,0),body.length);
    assert.equal(result.source.sheet,d.sheet);
  }
});

test('adapter reports missing resources and truncated row payload without returning a partial dataset',async()=>{
  globalThis.Excel={run:async()=>{throw Object.assign(new Error('private file identifier'),{code:'ItemNotFound'});}};
  await assert.rejects(readDataset(),e=>e.category==='missing_target'&&!e.message.includes('private'));
  const sheet={visibility:'Visible',id:'s',name:'Data',load(){},
    getRange(){return {load(){},values:[['A',1]],valueTypes:[['String','Double']]};}};
  globalThis.Excel={run:work=>work({sync:async()=>{},workbook:{getSelectedRange:()=>({worksheet:sheet,load(){},
    rowIndex:0,columnIndex:0,rowCount:4,columnCount:2})}})};
  await assert.rejects(readDataset(),e=>e.category==='partial_read');
  assert.equal(workbookFailure({code:'InvalidObjectPath'},'read').category,'stale_target');
});

test('table read uses current table geometry, excludes a native totals row, and preserves its identity',async()=>{
  const table={name:'Orders',showHeaders:true,showTotals:true,load(){},getRange(){return {
    rowIndex:4,columnIndex:2,rowCount:4,columnCount:2,load(){},
  };}};
  const read=[];
  const sheet={visibility:'Visible',id:'sheet-id',name:'Renamed',load(){},tables:{getItemOrNullObject:()=>table},
    getRange(address){read.push(address);return {load(){},rowIndex:4,columnIndex:2,rowCount:3,columnCount:2,
      values:[['Group','Amount'],['A',2],['B',3]],valueTypes:[['String','String'],['String','Double'],['String','Double']]};}};
  globalThis.Excel={run:work=>work({sync:async()=>{},workbook:{worksheets:{getItemOrNullObject:()=>sheet}}})};
  const data=await readDataset(undefined,undefined,{sheetId:'sheet-id',table:'Orders'});
  assert.equal(data.table,'Orders');assert.deepEqual(read,['C5:D7','C5:D7']);
  assert.equal(data.sheet,'Renamed');assert.equal(data.values.length,3);
  await assert.rejects(readDataset(undefined,undefined,{sheetId:'sheet-id',used:true,address:'A1:B3'}),/exactly one/);
});

test('diagnostics reports a host failure separately and contains no workbook values even when checks pass',async()=>{
  const fail=await runLightDiagnostics({inspectWorkbookStructure:async()=>{
    throw new WorkbookError('missing_target','enumerate','SECRET workbook name');
  }},100);
  assert.equal(fail[0].category,'missing_target');assert(!JSON.stringify(fail).includes('SECRET'));
  const office={supported:version=>version!=='1.13',
    inspectWorkbookStructure:async()=>({active:{id:'id',name:'PRIVATE'},sheets:[{id:'id',name:'PRIVATE'}]}),
    selectionAnchor:async()=>({id:'id',address:'A1:B2'}),
    readContext:async({address})=>{
      const wide=address.includes(':');const matrix=wide?[['PRIVATE','No'],['Text',3]]:[['PRIVATE']];
      return {values:matrix,formulas:matrix,types:matrix,formats:matrix};
    },inspectActiveObjects:async()=>({used:{address:'A1:B2'},tableCount:0,workbookNameCount:null})};
  const requests=[];
  const report=await runDiagnostics({office,probe:true,request:async(path,options)=>{
    requests.push([path,options.body]);
    return {ok:true,status:200,json:async()=>options.method?{status:'failed',category:'invalid_tool_sequence',private:'PRIVATE'}:
      path==='/health'?{status:'ok',build:'123abc9'}:{configured:true}};
  }});
  const json=JSON.stringify(report);
  assert(!json.includes('PRIVATE'));assert(!json.includes('A1:B2'));
  assert.equal(report.build,'123abc9');
  assert(report.checks.some(x=>x.category==='invalid_tool_sequence'&&x.status==='failed'));
  assert(report.checks.some(x=>x.operation==='Preview and write execution'&&x.status==='not_tested'));
  assert.deepEqual(requests.find(([,body])=>body),['/api/diagnostics','{}']);
});

test('the live-host evidence gate rejects a missing host, a failed scenario and a different build',()=>{
  const build='f106f6e56a18';
  const reports=['Excel Web','Excel Windows','Excel Mac'].map(host=>({
    tester:'fixture',
    diagnostics:{schema:1,build,host,reference:'fixture-reference',date:new Date().toISOString(),
      checks:[{category:'cell_readable',status:'passed'}]},
    results:scenarios.map(id=>({id,status:'passed',expected:'synthetic expected',observed:'synthetic observed'})),
  }));
  assert.deepEqual(verifyLiveReports(build,reports),[]);
  assert.match(verifyLiveReports(build,reports.slice(0,2))[0],/Mac/);
  reports[0].results[0].status='failed';
  assert.match(verifyLiveReports(build,reports)[0],/connect/);
  assert.match(verifyLiveReports('bad000000abc',reports)[0],/build metadata/);
});

test('workbook-scoped names must refer to an available range on the requested sheet',async()=>{
  globalThis.Office={context:{requirements:{isSetSupported:()=>true}}};
  const owner={id:'s',load(){}};
  const namedRange={worksheet:owner,isNullObject:false,rowIndex:1,columnIndex:3,rowCount:2,columnCount:2,load(){}};
  const name={type:'Range',load(){},getRangeOrNullObject:()=>namedRange};
  const sheet={id:'s',name:'Data',visibility:'Visible',load(){},
    getRange:()=>({load(){},values:[['Group','Sales'],['East',4]],valueTypes:[['String','String'],['String','Double']]})};
  globalThis.Excel={run:work=>work({sync:async()=>{},workbook:{worksheets:{getItemOrNullObject:()=>sheet},
    names:{getItemOrNullObject:()=>name}}})};
  const target={sheetId:'s',namedRange:'SalesInput'};
  assert.equal((await readDataset(undefined,undefined,target)).address,'D2:E3');
  owner.id='another';
  await assert.rejects(readDataset(undefined,undefined,target),error=>error.category==='invalid_reference');
  namedRange.isNullObject=true;
  await assert.rejects(readDataset(undefined,undefined,target),error=>error.category==='invalid_reference');
});

test('global resources bind to the owning worksheet and never fall back to the active selection',async()=>{
  globalThis.Office={context:{requirements:{isSetSupported:()=>true}}};
  const values=[['Class','Weight'],['Amber',5]];
  const dims={rowIndex:3,columnIndex:2,rowCount:2,columnCount:2,load(){}};
  const sheet={id:'other-sheet',name:'Other worksheet',visibility:'Visible',load(){},
    getRange(address){assert.equal(address,'C4:D5');
      return {...dims,values,valueTypes:[['String','String'],['String','Double']]};},
    getUsedRangeOrNullObject:()=>dims};
  const table={name:'OtherTable',showHeaders:true,showTotals:true,worksheet:sheet,load(){},
    getRange:()=>({...dims,rowCount:3})};
  const name={type:'Range',load(){},getRangeOrNullObject:()=>({...dims,worksheet:sheet})};
  const book={tables:{getItemOrNullObject:value=>value==='OtherTable'?table:{isNullObject:true,load(){}}},
    names:{getItemOrNullObject:()=>name},worksheets:{items:[sheet],load(){}},
    getSelectedRange(){assert.fail('Scope must never fall back to the selection');}};
  globalThis.Excel={run:work=>work({sync:async()=>{},workbook:book})};
  for (const target of [{table:'OtherTable'},{namedRange:'MyRange'},{worksheet:'oTHER WORKSHEET'}]) {
    const result=await readDataset(undefined,undefined,target,false);
    assert.equal(result.sheetId,'other-sheet');assert.equal(result.address,'C4:D5');
  }
  await assert.rejects(readDataset(undefined,undefined,{table:'Removed'}),e=>e.category==='missing_target');
  sheet.visibility='Hidden';
  await assert.rejects(readDataset(undefined,undefined,{table:'OtherTable'}),e=>e.category==='hidden_target');
  await assert.rejects(readDataset(undefined,undefined,{worksheet:'Other worksheet'}),e=>e.category==='hidden_target');
});

test('evidence reads require complete aligned metadata and retain computed formulas for lineage',async()=>{
  const block={load(){},values:[['Result'],[4]],valueTypes:[['String'],['Double']],
    text:[['Result'],['4']],numberFormat:[['General'],['General']],formulas:[['Result'],['=2+2']]};
  const sheet={id:'s',name:'Lab',visibility:'Visible',load(){},getRange:()=>block};
  globalThis.Excel={run:work=>work({sync:async()=>{},workbook:{getSelectedRange:()=>({worksheet:sheet,load(){},
    rowIndex:3,columnIndex:2,rowCount:2,columnCount:1})}})};
  const read=()=>readDataset(undefined,undefined,undefined,false,true);
  assert.deepEqual((await read()).formulas,[[null],['=2+2']]);
  for (const field of ['values','valueTypes','text','numberFormat','formulas']) {
    const original=block[field];block[field]=[original[0],[]];
    await assert.rejects(read(),e=>e.category==='partial_read');
    block[field]=original;
  }
});
