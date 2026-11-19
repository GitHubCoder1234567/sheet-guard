import test from 'node:test';
import assert from 'node:assert/strict';
import {a1,rect,audit,chartData,validate,conflict,correction} from '../public/engine.mjs';

const matrix=(v,callback)=>v.map((row,r)=>row.map((x,c)=>callback(x,r,c)));
function sheet(values,{formulas=values,r1c1=formulas,formats=matrix(values,()=> 'General'),
  types=matrix(values,x=>typeof x==='number'?'Double':x===''?'Empty':'String'),
  chartCount=0,partial=false,row=0,column=0}={}) {
  return {id:'sheet',name:'Sheet A',values,formulas,r1c1,formats,types,chartCount,partial,row,column};
}
const months=[['Month','Revenue'],['January',12],['February',19],['March',17],['April',25]];
const categories=[['Product','Sales'],['Laptops',12],['Audio',19],['Cameras',17]];
const action=(kind,address,fields={})=>({kind,address,sheetId:'sheet',...fields});

test('A1 addresses preserve offsets and restrict Excel bounds and cell limits',()=>{
  assert.equal(a1(4,27,2,2),'AB5:AC6');
  assert.deepEqual(rect('AB5:AC6'),{r:4,c:27,rows:2,columns:2});
  for(const invalid of ['A0','B1:A1','XFE1','A1:A1048577','A1:A2401','[Book]A1','X!A1','A1,C2','A1\n'])
    assert.throws(()=>rect(invalid));
});
test('repair uses translated consensus and distinguishes hardcodes and missing entries',()=>{
  for (const original of ['=A4+B4',12,'']) {
    const values=[['Unit','Price','Sales'],[2,3,6],[2,3,6],[2,3,original===''?'':5],[2,3,6],[2,3,6]];
    const formulas=matrix(values,(x,r,c)=>r&&c===2?(r===3?original:'=A'+(r+1)+'*B'+(r+1)):x);
    const r1c1=matrix(formulas,(v,r,c)=>r&&c===2?(r===3?(typeof original==='string'&&original.startsWith('=')?'=RC[-2]+RC[-1]':original):'=RC[-2]*RC[-1]'):v);
    const issues=audit([sheet(values,{formulas,r1c1})]);
    assert.equal(issues.filter(f=>f.action?.kind==='formula').length,1);
    assert.equal(issues.find(f=>f.action?.kind==='formula').action.address,'C4');
  }
});
test('mixed neighbouring formulas and text beginning with = produce no repair',()=>{
  const values=[['N','P','T'],[1,2,2],[1,2,2],[1,2,3],[1,2,'=RC[-2]*RC[-1]'],[1,2,2]];
  const formulas=matrix(values,(v,r,c)=>r&&c===2&&r!==4?'=A'+(r+1)+'*B'+(r+1):v);
  const r1c1=matrix(formulas,(v,r,c)=>r&&c===2?'=RC[-2]*RC[-1]':v);
  assert(!audit([sheet(values,{formulas,r1c1})]).some(f=>f.action?.kind==='formula'));
});
test('grouped profiling flags actual errors, numeric text, duplicate records, blanks and invalid dates',()=>{
  const values=[['Category','Units','Date','Price'],
    ['Audio',3,'2026-02-28',4],['audio ','3','31/02/2026',2],
    ['Other',4,'',2],['Other',4,'',2],['Spare',3,'02/03/2026','#REF!']];
  const types=matrix(values,x=>x==='#REF!'?'Error':typeof x==='number'?'Double':'String');
  const formats=matrix(values,(_,r,c)=>c===3?(r===1?'AUD':'USD'):'General');
  const results=audit([sheet(values,{types,formats})]);
  assert.equal(results[0].title,'Excel error values');
  for (const label of ['Duplicate','Numeric text','Whitespace','Missing','Invalid dates','Mixed number','Text variants']) {
    assert(results.some(f=>f.title.toLowerCase().includes(label.toLowerCase())),label);
  }
  assert.match(results.find(f=>f.title.startsWith('Invalid dates')).detail,/C3/);
  assert.doesNotMatch(results.find(f=>f.title.startsWith('Invalid dates')).detail,/C6/);
});
test('detects cycles in simple local arithmetic, including absolute references',()=>{
  const values=[['A','B','C'],[0,0,3],[1,1,3]];
  const formulas=[values[0],['=$B$2+C2','=A2',3],['=B3',1,3]];
  const issue=audit([sheet(values,{formulas})]).find(f=>f.title==='Circular formula references');
  assert.match(issue.detail,/A2, B2/);
  assert.doesNotMatch(issue.detail,/A3/);
});
test('trend and comparison suggestions require relevant, complete data and suppress existing charts',()=>{
  const trend=audit([sheet(months)]);
  assert.equal(trend.find(f=>f.category==='Visualisations').action.type,'line');
  assert.equal(trend.find(f=>f.category==='Analysis').action.kind,'dashboard');
  assert.equal(audit([sheet(categories)]).find(f=>f.category==='Visualisations').action.type,'bar');
  for (const s of [sheet(months,{chartCount:1}),sheet(months,{partial:true}),sheet([['Month','Revenue'],['Blue',1],['Red',5],['Green',3]])])
    assert(!audit([s]).some(f=>f.category==='Visualisations'));
});
test('chart types validate their input and destination bounds',()=>{
  assert.equal(chartData(sheet(months),'line').points[2][1],17);
  assert.throws(()=>chartData(sheet(months),'scatter'));
  assert.throws(()=>chartData(sheet([categories[0],['A',-2],['B',2],['C',4]]),'pie'));
  assert.throws(()=>chartData(sheet([categories[0],['A',3],['A',2],['B',4]]),'doughnut'));
  validate(action('chart','A1:B5',{type:'bar',title:'Sales',legend:false,position:'E3:M20'}));
  assert.throws(()=>validate(action('chart','A1:B5',{type:'PivotChart',title:'Sales',legend:false})));
});
test('unsupported actions, arbitrary formula payloads and ambiguous formats fail closed',()=>{
  assert(correction('=RC[-2]*RC[-1]',2));
  for(const expr of ['=WEBSERVICE("https://example.org")','=RC[0]*RC[-1]','=RC[-2]*RC[-1]+1','=RC[-999]*RC[-1]'])
    assert(!correction(expr,2),expr);
  for(const payload of [action('script','A1',{value:'alert(1)'}),action('formula','A1:A3',{value:'=RC[1]*RC[2]'}),
    action('number','B2',{value:'[$EUR]'}),action('heatmap','A1:B4')])
    assert.throws(()=>validate(payload));
});
test('bulk conflict detects overlapping rectangles, permits adjacent or distinct sheets',()=>{
  assert(conflict(action('number','B2:C4'),action('style','C4:D7')));
  assert(!conflict(action('number','B2:C4'),action('style','D4:D7')));
  assert(!conflict(action('number','B2:C4'),action('style','C4:D7',{sheetId:'other'})));
});

test('format standardisation is actionable for dominant date and percentage formats only',()=>{
  const values=[['Date','Discount','Price'],...Array.from({length:6},(_,i)=>[46000+i,.1,300+i])];
  const formats=matrix(values,(_,r,c)=>c===0?(r===5?'dd/mm/yyyy':'yyyy-mm-dd'):
    c===1?(r===5?'0.00':'0.0%'):(r===5?'USD':'AUD'));
  const f=audit([sheet(values,{formats})]).filter(f=>f.category==='Formatting');
  assert.equal(f.find(x=>x.title.includes('Date')).action.value,'yyyy-mm-dd');
  assert.equal(f.find(x=>x.title.includes('Discount')).action.value,'0.0%');
  assert.equal(f.find(x=>x.title.includes('Price')).action,undefined);
});

test('outlier messages identify the sampled cell with a numerical fence violation',()=>{
  const values=[['Spend'],...[1,2,3,2,4,5,4,3,100].map(x=>[x])];
  assert.match(audit([sheet(values)]).find(f=>f.title.includes('outliers')).detail,/A10/);
});

test('formula lexer covers A1 references in common calls without treating text and table fields as references',async()=>{
  const {formulaReferences}=await import('../public/engine.mjs');
  const refs=formulaReferences('=SUMIFS(\'Sales Data\'!$B$2:$B$12,Orders[A1],A2)+XLOOKUP(D2,Lookup!A1:A10,Lookup!B1:B10)+"Q9"');
  assert.deepEqual(refs.map(({sheet,first,last,absolute})=>[sheet,first,last,absolute]),[
    ['Sales Data','B2','B12',true],[undefined,'A2',undefined,false],[undefined,'D2',undefined,false],
    ['Lookup','A1','A10',false],['Lookup','B1','B10',false],
  ]);
});

test('advanced formula patterns and cross-sheet errors offer inspection without proposing a rewrite',()=>{
  const values=[['Input','Total'],...Array.from({length:7},(_,i)=>[i,100+i])];
  const formulas=matrix(values,(v,r,c)=>r&&c===1?(r===4?"=IF(A5,Other!B2,0)":'=IF(A'+(r+1)+',100,0)'):v);
  const r1c1=matrix(formulas,(v,r,c)=>r&&c===1?(r===4?'=IF(RC[-1],Other!R2C2,0)':'=IF(RC[-1],100,0)'):v);
  const source=sheet(values,{formulas,r1c1});
  const other={...sheet([['Input','Total'],['A','#REF!']],{types:[['String','String'],['String','Error']]}),id:'other',name:'Other'};
  const findings=audit([source,other]);
  assert(findings.some(f=>f.title.includes('Formula pattern exception') && !f.action && f.address==='B5'));
  assert(findings.some(f=>f.title.includes('Referenced sheet contains errors') && !f.action));
});

test('clarification requires user-confirmed metric and trim actions reject ambiguous text',async()=>{
  const {additiveMeasure,trimText}=await import('../public/engine.mjs');
  const values=[['Channel','Contribution'],['Retail',100],['Online',200],['Wholesale',300]];
  const item=audit([sheet(values)]).find(f=>f.clarification);
  assert.equal(item.clarification.metric,'Contribution');
  assert(!item.action);
  assert(!additiveMeasure('Contribution'));
  assert(additiveMeasure('Contribution',{confirmedMetric:'Contribution'}));
  assert(!additiveMeasure('Other',{confirmedMetric:'Contribution'}));
  assert.equal(trimText('  Audio  '),'Audio');
  for (const unsafe of [' =2+3 ',' -10 ',' 13 Oct ',' +A1 ']) assert.equal(trimText(unsafe),null);
});

test('format comparisons use actual source values and keep underlying numbers unchanged',async()=>{
  const {previewRows}=await import('../public/engine.mjs');
  const snapshot={...sheet([[.125]]),display:[['0.125']]};
  const [row]=previewRows({snapshot,action:{kind:'number',value:'0.0%'}});
  assert.equal(row.value,.125);
  assert.equal(row.before,'0.125');
  assert.match(row.after,/12[.,]5\s?%/);
  assert.equal(snapshot.values[0][0],.125);
});

test('keyed comparison aligns reordered records and excludes ambiguous IDs',async()=>{
  const {compareDatasets}=await import('../public/engine.mjs');
  const first={values:[['ID','Revenue','Old'],[1,20,'a'],[2,9,'x'],[2,10,'x'],[3,3,'f'],[4,0,'g'],['',8,'a']]};
  const second={values:[['Revenue','ID','New'],[5,4,'g'],[30,1,'z'],[17,2,'x'],[15,5,'y']]};
  const result=compareDatasets(first,second,'ID');
  assert.deepEqual([result.added,result.removed,result.changed,result.ambiguous,result.missing],[1,1,2,1,1]);
  assert.deepEqual(result.firstOnly,['Old']);
  assert.deepEqual(result.secondOnly,['New']);
  assert.equal(result.samples.find(s=>s.key===1).changes[0].percent,.5);
  assert.equal(result.samples.find(s=>s.key===4).changes[0].percent,undefined);
});

test('grouped analysis handles 30,000 rows and omits error and nonnumeric values',async t=>{
  const {summariseBy}=await import('../public/engine.mjs');
  const before=process.memoryUsage().heapUsed,start=performance.now();
  const values=[['Market','Sales'],...Array.from({length:30_000},(_,i)=>[i%2?'West':'East',10]),['West','invalid'],['#REF!',12]];
  const errors=Array(values.length).fill(null);errors[values.length-1]=[0];
  const result=summariseBy({values,errors},'Sales','Market');
  assert.equal(result.count,30_000);assert.equal(result.omitted,2);
  assert.deepEqual(result.groups.map(g=>g.sum),[150_000,150_000]);
  t.diagnostic('30,002 synthetic rows in '+(performance.now()-start).toFixed(1)+' ms; process heap delta '+((process.memoryUsage().heapUsed-before)/2**20).toFixed(1)+' MiB');
});

test('PivotTable preview rejects incomplete measures and Excel-style case collisions',async()=>{
  const {pivotPreview}=await import('../public/engine.mjs');
  const values=[['Group','Sales'],['East',1],['West',2],['East',3]];
  const action={dimension:'Group',measure:'Sales'};
  const preview=pivotPreview(sheet(values),action);
  assert.equal(preview.groups[0].sum,4);
  assert.throws(()=>pivotPreview(sheet([values[0],['East',1],['east',2]]),action),/clean text groups/);
  assert.throws(()=>pivotPreview(sheet([values[0],['East','1']]),action),/complete numeric/);
});

test('identifier matches distinguish different records from full-row duplicates',()=>{
  const result=audit([sheet([['Order ID','Value'],['A1',7],['A1',8],['B1',4],['B1',4]])]);
  const ids=result.filter(x=>x.title.startsWith('Repeated identifier'));
  assert.equal(ids.length,2);
  assert(ids.some(x=>x.detail.includes('differ') && x.matchRows.length===2 && !x.action));
  assert(ids.some(x=>x.detail.includes('match across')));
});
