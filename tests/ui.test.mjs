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
    getRange(address){return address==='D1:L17'||address==='D1'?slot:range(address);}};
  function range(address='A1:B5') {
    if (typeof address!=='string') address='A1:B5';
    const [,x,y,xx=x,yy=y]=/^([A-Z])(\\d+)(?::([A-Z])(\\d+))?$/.exec(address);
    const rowIndex=Number(y)-1,columnIndex=x.charCodeAt(0)-65,rowCount=Number(yy)-rowIndex,columnCount=xx.charCodeAt(0)-64-columnIndex;
    const slice=matrix=>matrix.slice(rowIndex,rowIndex+rowCount).map(row=>row.slice(columnIndex,columnIndex+columnCount));
    return {worksheet:sheet,values:slice(values),formulas:slice(values),formulasR1C1:slice(values),
      text:slice(values).map(row=>row.map(String)),
      valueTypes:slice(values).map(row=>row.map(x=>typeof x==='number'?'Double':'String')),
      get numberFormat(){return slice(formats);},set numberFormat(value){formats=value;calls.push('number format write');},
      rowIndex,columnIndex,rowCount,columnCount,
      rowHidden:false,columnHidden:false,hasSpill:false,
      select(){window.sheetGuardSelection='A1:B5';},getColumn(){return this;},getResizedRange(){return this;},getOffsetRange(){return this;},
      getMergedAreasOrNullObject(){return {isNullObject:true,load(){}};},
      load(){},isNullObject:false};
  }
  const slot = {left:100,top:0,width:400,height:200,load(){},
    getUsedRangeOrNullObject(){return {isNullObject:true,load(){}};}};
  const book = {protection:{protected:false,load(){}},worksheets:{items:[sheet],load(){},getItem(){return sheet;},getActiveWorksheet(){return sheet;}},getSelectedRange:range};
  window.Office={HostType:{Excel:'Excel'},onReady:(fn)=>Promise.resolve(fn({host:'Excel'})),
    context:{requirements:{isSetSupported:()=>true}}};
  window.Excel={run:fn=>fn({workbook:book,sync:async()=>{}})};
`;

async function openView(page,name) {
  await page.locator('summary[aria-label="Workspace menu"]').click();
  await page.getByRole('button',{name,exact:true}).click();
}

async function fixture(t, office=workbookDouble, agent=null) {
  const browser=await chromium.launch({executablePath:chromiumPath,headless:true});
  t.after(()=>browser.close());
  const page=await browser.newPage({viewport:{width:360,height:760}});
  page.setDefaultTimeout(7000);
  const errors=[];
  page.on('pageerror',e=>errors.push(e.message));
  await page.route(origin+'/**',async route=>{
    const path=new URL(route.request().url()).pathname;
    if (path==='/health' && !hostedOrigin) return route.fulfill({json:{status:'ok',build:'fa1f913'}});
    if (path==='/api/diagnostics' && !hostedOrigin) return route.fulfill({json:route.request().method()==='POST'?
      {status:'passed',category:'structured_greeting'}:{configured:true,provider:'groq'}});
    if (path==='/api/agent' && agent) {
      const result=agent(route.request().postDataJSON());
      return route.fulfill(result.status?result:{json:result});
    }
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


test('AI routes a natural-language ranking through the local aggregation tool',async t=>{
  let hits=0;
  const {page,errors}=await fixture(t,workbookDouble,request=>{
    hits++;
    assert.match(request.question,/rank the months/i);
    return {reply:'Requesting totals.',intent:'none',option:'default',finding:-1,tool:'AggregateData',
      sheet:'',address:'',measure:'Revenue',dimension:'Month',operation:'highest'};
  });
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  await page.getByRole('textbox',{name:'Workbook question'}).fill('Rank the months by highest revenue');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('Highest total Revenue: April · 26',{exact:false}).waitFor();
  assert.equal(hits,1);
  assert.deepEqual(errors,[]);
});


test('read-only workbook questions answer with local evidence without sharing or writes',async t=>{
  const mock=workbookDouble.replace("['April',26]", "['April','#DIV/0!']").replace(
    "typeof x==='number'?'Double':'String'", "x==='#DIV/0!'?'Error':typeof x==='number'?'Double':'String'");
  let requests=0;
  const {page,errors}=await fixture(t,mock,()=>{requests++;return {status:500};});
  for (const [question,answer] of [
    ['What is the current worksheet called?','The current worksheet is “Monthly”.'],
    ['How many records are in this dataset?','4 records in Monthly!A1:B5'],
    ['What columns are in this dataset?','2 columns in Monthly!A1:B5: Month, Revenue.'],
    ['Which month has the highest revenue?','Highest total Revenue: March · 22'],
    ['Find formula errors','Found 1 Excel error values in Monthly!A1:B5: B5 #DIV/0!'],
  ]) {
    await page.getByRole('textbox',{name:'Workbook question'}).fill(question);
    await page.getByRole('button',{name:'Send',exact:true}).click();
    await page.getByText(answer,{exact:false}).waitFor();
  }
  assert.equal(requests,0);
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});

test('chat retains the failed draft and displays actual authentication and malformed-response errors',async t=>{
  let hits=0;
  const {page}=await fixture(t,workbookDouble,()=>{
    hits++;
    if (hits===1) return {status:401,json:{error:'The deployment access password is invalid.'}};
    if (hits===2) return {status:200,body:'unexpected HTML',contentType:'text/html'};
    return {reply:'Hello. I can help review your workbook.',intent:'none',option:'default',finding:-1};
  });
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  const question=page.getByRole('textbox',{name:'Workbook question'});
  await question.fill('Hello Forma');
  const send=page.getByRole('button',{name:'Send',exact:true});
  await send.click();
  await page.getByRole('alert').getByText('The deployment access password is invalid.').waitFor();
  assert.equal(await question.inputValue(),'Hello Forma');
  await send.click();
  await page.getByRole('alert').getByText('The AI service returned an invalid answer.',{exact:false}).waitFor();
  await send.click();
  await page.getByText('Hello. I can help review your workbook.').waitFor();
  assert.equal(await question.inputValue(),'');
  assert.equal(await page.locator('article').getByText('Hello Forma',{exact:true}).count(),1);
});

test('formula explanation uses actual Excel formula and result without a write',async t=>{
  const mock=workbookDouble.replace('getSelectedRange:range', 'getSelectedRange:formulaCell')+
    `function formulaCell(){return {...range(),rowIndex:1,columnIndex:1,rowCount:1,columnCount:1,
      values:[[12]],formulas:[['=A2*3']],formulasR1C1:[['=RC[-1]*3']],numberFormat:[['General']],
      text:[['12']],valueTypes:[['Double']]};}`;
  const {page}=await fixture(t,mock);
  await page.getByRole('textbox',{name:'Workbook question'}).fill('Explain this formula');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('Current result: 12',{exact:false}).waitFor();
  await page.getByText('References: A2',{exact:false}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
});

test('tool evidence returns an answer, and a later AI recommendation still requires approval',async t=>{
  const {page}=await fixture(t,workbookDouble,payload=>{
    const base={reply:'Checking workbook context.',intent:'none',option:'default',finding:-1,
      sheet:'',address:'',measure:'',dimension:'',operation:''};
    if (payload.toolResult) {
      assert.equal(payload.toolResult.worksheets[0].name,'Monthly');
      return {...base,tool:'none',reply:'I found the Monthly worksheet.'};
    }
    return payload.question==='What is in my workbook?'?{...base,tool:'InspectWorkbook'}:
      {...base,tool:'GenerateChartProposal',option:'line'};
  });
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  const question=page.getByRole('textbox',{name:'Workbook question'});
  const send=page.getByRole('button',{name:'Send',exact:true});
  await question.fill('What is in my workbook?');await send.click();
  await page.getByText('I found the Monthly worksheet.').waitFor();
  await question.fill('Please suggest a chart.');await send.click();
  await page.getByRole('button',{name:'Review proposal'}).click();
  const card=page.locator('article').filter({has:page.getByRole('button',{name:'Create chart',exact:true})});
  await card.getByRole('button',{name:'Preview',exact:true}).click();
  await card.getByRole('button',{name:'Apply changes',exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
});

test('exact multi-part sales question gives a ranked table and source cells instead of opening audit',async t=>{
  let requests=0;
  const mock=workbookDouble.replace("['Month','Revenue'],['January',12],['February',18],['March',22],['April',26]",
    "['Category','Sales'],['Office',12],['Travel',18],['Office',22],['Meals',26]");
  const {page}=await fixture(t,mock,()=>{requests++;return {status:500};});
  await page.getByRole('textbox',{name:'Workbook question'}).fill('Analyse the current worksheet and identify which product category generates the highest total sales. Calculate the total sales for every category, rank them, explain the calculations, and identify the source cells.');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('Highest total Sales: Office · 34',{exact:false}).waitFor();
  await page.getByText('Monthly!B2:B5 for each matching value in Monthly!A2:A5',{exact:false}).waitFor();
  assert.deepEqual(await page.locator('.ranking-table tbody tr').allTextContents(),['1Office342','2Meals261','3Travel181']);
  assert.equal(await page.getByRole('heading',{name:'Workbook review'}).count(),0);
  assert.equal(requests,0);
});

test('cloud calculation can inspect and then aggregate; audit plan cannot replace its answer',async t=>{
  const base={reply:'Checking workbook context.',intent:'none',option:'default',finding:-1,sheet:'',address:'',measure:'',dimension:'',operation:''};
  let calls=0;
  const {page}=await fixture(t,workbookDouble,payload=>{
    calls++;
    if (payload.toolResult) {
      assert.equal(payload.toolResult.calculationPending,true);
      return {...base,tool:'AggregateData',measure:'Revenue',dimension:'Month',operation:'total'};
    }
    return payload.question.includes('bad route')?{...base,tool:'none',intent:'audit'}:{...base,tool:'InspectRange'};
  });
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  const q=page.getByRole('textbox',{name:'Workbook question'}), send=page.getByRole('button',{name:'Send',exact:true});
  await q.fill('Calculate grouped revenue from the selection and explain the source.');
  await send.click();
  await page.locator('.ranking-table tbody tr').first().waitFor();
  assert.equal(calls,2);
  await q.fill('Calculate grouped revenue from the selection, bad route');
  await send.click();
  await page.getByRole('alert').getByText('No calculation was performed.',{exact:false}).waitFor();
  assert.equal(await page.getByRole('heading',{name:'Workbook review'}).count(),0);
});

test('cleaning keeps the cleaning view, while unsupported statistics explain the limitation',async t=>{
  const {page}=await fixture(t);
  const q=page.getByRole('textbox',{name:'Workbook question'}), send=page.getByRole('button',{name:'Send',exact:true});
  await q.fill('Clean this dataset');await send.click();
  await page.getByRole('heading',{name:'Cleaning opportunities'}).waitFor();
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  await q.fill('Calculate median revenue in this worksheet');await send.click();
  await page.getByRole('alert').getByText('Forma cannot verify that statistic',{exact:false}).waitFor();
});

test('Sales Audit chat carries the Electronics ledger through variants, preview, approval, undo and stale evidence',async t=>{
  const data=JSON.parse(await readFile(new URL('./fixtures/sales-audit.json',import.meta.url)));
  const excel=`
    const data=${JSON.stringify(data)}, writes=window.auditWrites=[];
    const noop=()=>{};
    function make(address='A1:H17'){
      const [,left,top,right=left,bottom=top]=/^([A-Z])(\\d+)(?::([A-Z])(\\d+))?$/.exec(address);
      const r=Number(top)-1,c=left.charCodeAt(0)-65,rows=Number(bottom)-r,columns=right.charCodeAt(0)-64-c;
      const subset=matrix=>matrix.slice(r,r+rows).map(row=>row.slice(c,c+columns));
      return {worksheet:sheet,rowIndex:r,columnIndex:c,rowCount:rows,columnCount:columns,
        load:noop,select(){window.auditSelected=address;},rowHidden:false,columnHidden:false,hasSpill:false,
        getMergedAreasOrNullObject(){return {isNullObject:true,load:noop};},
        getTables(){return {load:noop,items:[]};},dataValidation:{load:noop,type:'None'},
        get values(){return subset(data.values);},
        set values(value){data.values[r][c]=value[0][0];data.display[r][c]=String(value[0][0]);writes.push([address,value[0][0]]);},
        get formulas(){return subset(data.values);},get formulasR1C1(){return subset(data.values);},
        get numberFormat(){return subset(data.formats);},get text(){return subset(data.display);},
        get valueTypes(){return subset(data.values.map((row,i)=>row.map((v,j)=>data.errors[i].includes(j)?'Error':typeof v==='number'?'Double':v===''?'Empty':'String')));}
      };
    }
    const sheet={name:data.sheet,id:data.sheetId,visibility:'Visible',load:noop,activate:noop,
      protection:{load:noop,protected:false},charts:{load:noop,items:[]},
      getUsedRangeOrNullObject(){return make();},getRange:make};
    const wb={getSelectedRange:()=>make(),worksheets:{load:noop,items:[sheet],getItem:()=>sheet,getActiveWorksheet:()=>sheet}};
    window.Office={HostType:{Excel:'Excel'},onReady:fn=>Promise.resolve(fn({host:'Excel'})),
      context:{requirements:{isSetSupported:()=>true}}};
    window.Excel={run:fn=>fn({workbook:wb,sync:async()=>{}})};
    window.changeSale=()=>{data.values[1][4]=4000;data.display[1][4]='AUD 4000';};
  `;
  let explanations=0;
  const {page,errors}=await fixture(t,excel,payload=>{
    explanations++;
    assert.deepEqual(Object.keys(payload).sort(),['question','toolResult']);
    assert.equal(payload.toolResult.verifiedCalculation.operation,'sum');
    assert(!JSON.stringify(payload).includes('SG-1001'));
    return explanations===1?{status:502,json:{error:'Groq reached its output limit.',code:'truncated',reference:'aabbccddeeff'}}:
      {reply:'The calculation adds eligible sales by exact category label. A blank sales cell is excluded.',
        intent:'none',option:'default',finding:-1,tool:'none'};
  });
  const say=async text=>{
    await page.getByRole('textbox',{name:'Workbook question'}).fill(text);
    await page.getByRole('button',{name:'Send',exact:true}).click();
  };
  await say('Which product category has the highest sales? Show the exact calculations and source rows.');
  await page.getByText('Highest total Sales: Electronics · 4,780',{exact:false}).waitFor();
  assert.equal(explanations,0);
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  const exact='Analyse the Sales Audit worksheet. Rank every product category by total sales. Show the exact cells used for each calculation.';
  await say(exact);
  await page.getByText('Groq reached its output limit.',{exact:false}).waitFor();
  assert.equal(explanations,1);
  await page.getByText('Highest total Sales: Electronics · 4,780',{exact:false}).last().waitFor();
  assert.equal(await page.getByRole('alert').count(),0);
  await say(exact);
  await page.getByText('AI explanation · The calculation adds eligible sales by exact category label.',{exact:false}).waitFor();
  assert.equal(explanations,2);
  await say('Where exactly did 4,780 come from?');
  await page.getByText('3400 + 1380 = 4780',{exact:false}).waitFor();
  await page.getByText('Sales Audit!B2 · Electronics / Sales Audit!E2 = 3400 · A2 SG-1001',{exact:true}).waitFor();
  await page.getByText('Sales Audit!B3 · Electronics / Sales Audit!E3 = 1380 · A3 SG-1002',{exact:true}).waitFor();
  await page.getByRole('button',{name:'Select E2',exact:true}).click();
  assert.equal(await page.evaluate(()=>window.auditSelected),'E2');
  await page.getByRole('button',{name:'Return to previous selection'}).click();
  assert.equal(await page.evaluate(()=>window.auditSelected),'A1:H17');
  await say('But there is also lowercase electronics.');
  await page.getByText('combined total is 6905',{exact:false}).waitFor();
  await page.getByText('Sales Audit!B4 · electronics / Sales Audit!E4 = 2125',{exact:false}).waitFor();
  await say('Fix the inconsistent category labels.');
  await page.getByRole('button',{name:'Review proposal'}).click();
  const card=page.locator('article').filter({has:page.getByRole('button',{name:'Standardise category label',exact:true})});
  await card.getByRole('button',{name:'Preview',exact:true}).click();
  await card.getByText('Before · "electronics"',{exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.auditWrites),[]);
  await card.getByRole('button',{name:'Apply changes',exact:true}).click();
  await page.getByText('Change verified in Excel.',{exact:true}).waitFor();
  await openView(page,'History');
  page.once('dialog',d=>d.accept());
  await page.getByRole('button',{name:'Undo verified change'}).click();
  await page.getByText('Undo verified in Excel.',{exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.auditWrites),[['B4','Electronics'],['B4','electronics']]);
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  await page.evaluate(()=>window.changeSale());
  await say('Show me the rows behind the total.');
  await page.getByRole('alert').getByText('The source cells changed',{exact:false}).waitFor();
  assert.deepEqual(errors,[]);
});

test('diagnostics separates actual reads, synthetic provider success and untested writes',async t=>{
  const {page}=await fixture(t);
  await page.getByText('Workbook readable',{exact:true}).waitFor();
  await openView(page,'Settings');
  await page.getByText('Diagnostics',{exact:true}).click();
  await page.getByLabel('Include live Groq check').check();
  await page.getByRole('button',{name:'Run diagnostics',exact:true}).click();
  await page.getByText('Single-cell read · Passed',{exact:true}).waitFor();
  await page.getByText('Groq connectivity and structured reply · Passed',{exact:true}).waitFor();
  await page.getByText('Preview and write execution · Not tested',{exact:true}).waitFor();
  await page.getByText('View safe report JSON',{exact:true}).click();
  const report=JSON.parse(await page.locator('details pre').textContent());
  const json=JSON.stringify(report);
  assert.equal(report.build,'fa1f913');
  for (const secret of ['January','Monthly','Revenue','A1:B5']) assert(!json.includes(secret),secret);
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
});

test('column median uses selected workbook values, exposes cells and never calls AI or writes',async t=>{
  const {page}=await fixture(t);
  let requests=0;
  page.on('request',request=>{if (request.url().endsWith('/api/agent')) requests++;});
  await page.getByRole('textbox',{name:'Workbook question'}).fill('Calculate the median of Revenue in the selection');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('median of Revenue · 20.',{exact:false}).waitFor();
  await page.getByText('Monthly!B2 = 12',{exact:false}).waitFor();
  await page.getByText('Monthly!B5 = 26',{exact:false}).waitFor();
  assert.equal(requests,0);
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
});

test('conditional ranking stays local and visibly rechecks original measure and predicate cells',async t=>{
  const mock=workbookDouble
    .replace("['Month','Revenue'],['January',12],['February',18],['March',22],['April',26]",
      "['Unit','Hours','Phase','Priority'],['Alpha',12,'Active',2],['Beta',90,'Closed',3],['Beta',21,'Active',4],['Alpha',17,'Active',3]")
    .replaceAll('A1:B5','A1:D5')
    .replace("let formats=", "window.changePriority=()=>{values[1][3]=5;}; let formats=");
  const cloud=[];
  const {page,errors}=await fixture(t,mock,payload=>{
    cloud.push(payload);
    return {reply:'Hello.',intent:'none',option:'default',finding:-1};
  });
  const q=page.getByRole('textbox',{name:'Workbook question'});
  const send=page.getByRole('button',{name:'Send',exact:true});
  await q.fill('Rank "Unit" by total "Hours" where "Phase" = "Active" and "Priority" >= 2');
  await send.click();
  await page.getByText('Highest total Hours: Alpha · 29',{exact:false}).waitFor();
  assert.deepEqual(await page.locator('.ranking-table tbody tr').allTextContents(),['1Alpha292','2Beta211']);
  await page.getByRole('button',{name:'Alpha',exact:true}).click();
  await page.getByText('Monthly!C2 (Phase) = "Active" · Monthly!D2 (Priority) = 2',{exact:false}).waitFor();
  await page.getByText('Monthly!A5 · Alpha / Monthly!B5 = 17',{exact:true}).waitFor();
  await page.evaluate(()=>window.changePriority());
  await page.getByRole('button',{name:'Alpha',exact:true}).click();
  await page.getByText('The source cells changed',{exact:false}).waitFor();
  assert.deepEqual(cloud,[]);
  await openView(page,'Settings');
  await page.getByRole('checkbox',{name:'Allow AI data sharing for this session'}).check();
  await page.getByRole('button',{name:'Back to workspace',exact:true}).click();
  await q.fill('Hello');await send.click();
  await page.getByText('Hello.',{exact:true}).waitFor();
  assert.equal(cloud.length,1);
  assert.deepEqual(cloud[0].conversation,[]);
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});

test('cross-sheet key inspection warns about record multiplication without writing or calling AI',async t=>{
  const office=workbookDouble.replace("const values = [['Month','Revenue'],['January',12],['February',18],['March',22],['April',26]];",
    "let values=[['Order key','Time'],['X',12],['X',18],['Y',22],['Left',26]];")
    .replace("const book = {","window.showOther=()=>{sheet.id='other-sheet';sheet.name='Shipments';"+
      "values=[['Dispatch key','Status'],['X','Open'],['X','Done'],['Y','Open'],['Right','Done']];}; const book = {");
  let requests=0;
  const {page,errors}=await fixture(t,office,()=>{requests++;return {status:500};});
  await page.getByRole('button',{name:'Compare datasets',exact:true}).click();
  await page.getByRole('button',{name:'Capture first selection',exact:true}).click();
  await page.getByText('Monthly!A1:B5 · 4 data rows',{exact:true}).waitFor();
  await page.evaluate(()=>window.showOther());
  await page.getByRole('button',{name:'Capture second selection',exact:true}).click();
  await page.getByText('Shipments!A1:B5 · 4 data rows',{exact:true}).waitFor();
  await page.getByText('Inspect key relationships',{exact:true}).click();
  await page.getByLabel('First key',{exact:true}).selectOption('Order key');
  await page.getByLabel('Second key',{exact:true}).selectOption('Dispatch key');
  await page.getByRole('button',{name:'Check relationships',exact:true}).click();
  await page.getByText('5 row pairs would match in an inner join.',{exact:true}).waitFor();
  await page.getByText('Many-to-many keys would multiply records.',{exact:false}).waitFor();
  await page.getByText('Sample key cells',{exact:true}).click();
  await page.getByText('"X" · Monthly!A2, Monthly!A3 ↔ Shipments!A2, Shipments!A3',{exact:true}).waitFor();
  assert.equal(requests,0);
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});

test('named-table plans use the table owner and surface deletion before showing saved evidence',async t=>{
  const office=workbookDouble.replace("const book = {",
    "const owner={...sheet,id:'remote-sheet',name:'Second worksheet'};"+
    "let deleted=false;window.deleteDataset=()=>{deleted=true;};"+
    "const dataset={name:'ImportedData',showHeaders:true,showTotals:false,worksheet:owner,load(){},getRange:()=>range()};"+
    "const book = {tables:{getItemOrNullObject:()=>deleted?{isNullObject:true,load(){}}:dataset},");
  let requests=0;
  const {page,errors}=await fixture(t,office,()=>{requests++;return {status:500};});
  await page.getByRole('textbox',{name:'Workbook question'}).fill(
    'Rank "Month" by total "Revenue" where "Revenue" >= 18 in table "ImportedData"');
  await page.getByRole('button',{name:'Send',exact:true}).click();
  await page.getByText('Highest total Revenue: April · 26',{exact:false}).waitFor();
  await page.getByRole('button',{name:'April',exact:true}).click();
  await page.getByText('Second worksheet!A5 · April / Second worksheet!B5 = 26',{exact:true}).waitFor();
  await page.evaluate(()=>window.deleteDataset());
  await page.getByRole('button',{name:'April',exact:true}).click();
  await page.getByText('The selected Excel table no longer exists.',{exact:true}).waitFor();
  assert.equal(requests,0);
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
  assert.deepEqual(errors,[]);
});

test('worksheet discovery reads fresh cells and reports the same worksheet identity as the sample',async t=>{
  const office=workbookDouble.replace('let formats=','window.renameField=()=>{values[0][1]="Utilization";};let formats=');
  const {page}=await fixture(t,office);
  await page.evaluate(()=>window.renameField());
  await openView(page,'Audit');
  await page.getByRole('button',{name:'Discover worksheet datasets',exact:true}).click();
  await page.getByRole('heading',{name:'Monthly',exact:true}).waitFor();
  await page.getByText('A1:B5 · supported',{exact:true}).click();
  await page.getByText('Utilization · numeric measure candidate',{exact:true}).waitFor();
  assert.deepEqual(await page.evaluate(()=>window.sheetGuardTestCalls),[]);
});
