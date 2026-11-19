import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {chromium} from 'playwright';

const chromiumPath=process.env.CHROMIUM_EXECUTABLE;
const hostedOrigin=process.env.SHEETGUARD_TEST_ORIGIN;
const origin=hostedOrigin || 'https://sheetguard.test';
const workbookDouble = `
  const values = [['Month','Revenue'],['January',12],['February',18],['March',22],['April',26]];
  let formats=values.map(row=>row.map(()=> 'General'));
  const calls = window.sheetGuardTestCalls=[];
  const charts={items:[],load(){},add(type) {
    calls.push('chart created');
    const load=()=>{};
    const series={load,setXAxisValues(r){calls.push('categories linked');},setValues(r){calls.push('values linked');}};
    const chart={load,name:'Chart 1',chartType:type,left:100,top:0,width:400,height:200,
      title:{load},legend:{load},axes:{categoryAxis:{title:{load}},valueAxis:{title:{load}}},
      series:{load,items:[series]},setPosition(){calls.push('chart positioned');}};
    this.items.push(chart);
    return chart;
  }};
  const sheet = {id:'excel-sheet',name:'Monthly',visibility:'Visible',load(){},activate(){},protection:{protected:false,load(){}},
    charts,getUsedRangeOrNullObject:()=>range(),
    getRange(address){return address==='D1:L17'||address==='D1'?slot:range();}};
  function range() {
    return {worksheet:sheet,values,formulas:values,formulasR1C1:values, text:values.map(row=>row.map(String)),
      valueTypes:values.map(row=>row.map(x=>typeof x==='number'?'Double':'String')),
      get numberFormat(){return formats;},set numberFormat(value){formats=value;calls.push('number format write');},rowIndex:0,columnIndex:0,rowCount:5,columnCount:2,
      rowHidden:false,columnHidden:false,hasSpill:false,
      select(){window.sheetGuardSelection='A1:B5';},getColumn(){return this;},getResizedRange(){return this;},getOffsetRange(){return this;},
      getMergedAreasOrNullObject(){return {isNullObject:true,load(){}};},
      load(){},isNullObject:false};
  }
  const slot = {left:100,top:0,width:400,height:200,load(){},
    getUsedRangeOrNullObject(){return {isNullObject:true,load(){}};}};
  const book = {protection:{protected:false,load(){}},worksheets:{items:[sheet],load(){},getItem(){return sheet;}},getSelectedRange:range};
  window.Office={HostType:{Excel:'Excel'},onReady:(fn)=>Promise.resolve(fn({host:'Excel'})),
    context:{requirements:{isSetSupported:()=>true}}};
  window.Excel={run:fn=>fn({workbook:book,sync:async()=>{}})};
`;

async function openView(page,name) {
  await page.locator('summary[aria-label="Workspace menu"]').click();
  await page.getByRole('button',{name,exact:true}).click();
}

async function fixture(t, office=workbookDouble) {
  const browser=await chromium.launch({executablePath:chromiumPath,headless:true});
  t.after(()=>browser.close());
  const page=await browser.newPage({viewport:{width:360,height:760}});
  page.setDefaultTimeout(7000);
  const errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  await page.route(origin+'/**',async route=>{
    const path=new URL(route.request().url()).pathname;
    if (path==='/api/agent') return route.fulfill({json:{reply:'Monthly revenue increased from 12 to 26.',intent:'none',option:'default',finding:-1}});
    if (hostedOrigin) return route.continue();
    const file=path==='/'?'index.html':path.slice(1);
    try {
      const content=await readFile(new URL('../public/'+file,import.meta.url));
      const type=file.endsWith('.mjs')?'text/javascript':file.endsWith('.css')?'text/css':file.endsWith('.png')?'image/png':'text/html';
      await route.fulfill({body:content,contentType:type});
    } catch {await route.fulfill({status:404});}
  });
  await page.route('https://appsforoffice.microsoft.com/**',route=>route.fulfill({body:office,contentType:'text/javascript'}));
  await page.goto(origin+'/index.html');
  await page.getByText('Audit complete.',{exact:true}).waitFor();
  return {page,errors};
}
test('audit presents a preview before approval and a declined proposal appears in history',async t=>{
  const {page,errors}=await fixture(t);
  await openView(page,'Recommendations');
  const recommendation=page.locator('article').filter({hasText:'Show the trend over time'});
  assert.equal(await recommendation.getByRole('button',{name:'Apply changes',exact:true}).count(),0);
  await recommendation.getByRole('button',{name:'Preview',exact:true}).click();
  await recommendation.getByRole('img',{name:'line chart of Revenue by Month'}).waitFor();
  assert.equal(await page.evaluate(()=>window.sheetGuardSelection),'A1:B5');
  await recommendation.getByRole('button',{name:'Apply changes',exact:true}).waitFor();
  if (process.env.SHEETGUARD_UI_SCREENSHOT) await page.screenshot({path:process.env.SHEETGUARD_UI_SCREENSHOT,fullPage:true});
  await recommendation.getByRole('button',{name:'Decline'}).click();
  assert.equal(await recommendation.count(),0);
  await openView(page,'History');
  await page.getByText('Declined · Monthly!A1:B5',{exact:true}).waitFor();
  assert.deepEqual(errors,[]);
});
test('chat requires consent and displays a validated backend reply',async t=>{
  const {page,errors}=await fixture(t);
  await page.getByRole('textbox',{name:'Workbook question'}).fill('How did revenue change?');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('Enable AI data sharing in Settings',{exact:false}).waitFor();
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  await page.getByRole('textbox',{name:'Workbook question'}).fill('How did revenue change?');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('Monthly revenue increased from 12 to 26.').waitFor();
  assert.deepEqual(errors,[]);
});

test('approved chart uses the chart API and verifies placement',async t=>{
  const {page,errors}=await fixture(t);
  await openView(page,'Visualisations');
  const card=page.locator('article').filter({hasText:'Show the trend over time'});
  await card.getByRole('button',{name:'Preview',exact:true}).click();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  await card.getByRole('button',{name:'Apply changes',exact:true}).click();
  await page.getByText('Change verified in Excel.',{exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),
    ['chart created','categories linked','values linked','chart positioned']);
  assert.deepEqual(errors,[]);
});

test('approved formatting writes through Office adapter and undo restores the prior format',async t=>{
  const {page,errors}=await fixture(t);
  await openView(page,'Recommendations');
  await page.getByText('Prepare a change from the selection',{exact:true}).click();
  await page.getByLabel('Operation').selectOption('number');
  await page.getByRole('button',{name:'Prepare proposal'}).click();
  const card=page.locator('article').filter({has:page.getByRole('heading',{name:'Change number display'})});
  await card.getByRole('button',{name:'Preview',exact:true}).click();
  await card.getByText('General',{exact:true}).first().waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  await card.getByRole('button',{name:'Apply changes',exact:true}).click();
  await page.getByText('Change verified in Excel.',{exact:true}).waitFor();
  await openView(page,'History');
  page.once('dialog',dialog=>dialog.accept());
  await page.getByRole('button',{name:'Undo verified change'}).click();
  await page.locator('article').getByText('Undo verified',{exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),['number format write','number format write']);
  assert.deepEqual(errors,[]);
});


test('observations can be inspected and dismissed without a workbook write',async t=>{
  const {page,errors}=await fixture(t,workbookDouble.replace("['April',26]","[' April ',26]"));
  await openView(page,'Recommendations');
  const card=page.locator('article').filter({has:page.getByRole('button',{name:'Leading or trailing whitespace',exact:true})});
  await card.getByRole('button',{name:'Inspect',exact:true}).click();
  await card.getByText('Workbook context',{exact:true}).waitFor();
  assert.equal(await card.getByRole('button',{name:'Apply changes',exact:true}).count(),0);
  await card.getByRole('button',{name:'Explain',exact:true}).click();
  await card.getByText('This finding does not contain a proposed workbook edit.',{exact:false}).waitFor();
  await card.getByRole('button',{name:'Dismiss',exact:true}).click();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});

test('dashboard clarification generates a KPI preview after explicit confirmation',async t=>{
  const {page,errors}=await fixture(t,workbookDouble.replace("'Revenue'","'Contribution'"));
  await openView(page,'Recommendations');
  const card=page.locator('article').filter({has:page.getByRole('button',{name:'Confirm how to summarise Contribution',exact:true})});
  await card.getByRole('button',{name:'Provide information',exact:true}).click();
  assert(await card.getByRole('button',{name:'Generate proposal',exact:true}).isDisabled());
  await card.getByLabel('Can Contribution be added across these rows?').selectOption('sum');
  await card.getByRole('button',{name:'Generate proposal',exact:true}).click();
  await card.getByText('78',{exact:true}).waitFor();
  await card.getByText('Observations',{exact:true}).waitFor();
  await card.getByRole('img',{name:'line chart of Contribution by Month'}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});

test('grouped totals are calculated from the selection without enabling AI',async t=>{
  const {page,errors}=await fixture(t);
  await page.getByRole('heading',{name:'Forma',exact:true}).waitFor();
  await page.getByRole('textbox',{name:'Workbook question'}).fill('total Revenue by Month');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('April: 26 (1 rows)',{exact:false}).waitFor();
  await page.getByText('Source: Monthly!A1:B5',{exact:false}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});
test('explicit PivotTable request produces a read-only proposal preview',async t=>{
  const {page,errors}=await fixture(t);
  await page.getByRole('textbox',{name:'Workbook question'}).fill('Create a pivot table for total Revenue by Month');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByRole('button',{name:'Review proposal',exact:true}).click();
  const card=page.locator('article').filter({hasText:'PivotTable · Revenue by Month'});
  await card.getByRole('button',{name:'Preview',exact:true}).click();
  await card.getByText('Native PivotTable · new worksheet',{exact:true}).waitFor();
  await card.getByRole('button',{name:'Apply changes',exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});
