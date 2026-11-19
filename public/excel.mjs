import {a1, rect, validate, chartData, chartTypes, pivotPreview, headingStyle, correction, formula, trimText, additiveMeasure, LIMIT, DATASET_CELL_LIMIT, DATASET_ROW_LIMIT} from './engine.mjs';

const properties = 'rowIndex,columnIndex,rowCount,columnCount,values,formulas,formulasR1C1,numberFormat,valueTypes,text';
export const supported = version => Office.context.requirements.isSetSupported('ExcelApi',version);
const equal = (a,b) => JSON.stringify(a) === JSON.stringify(b);
const read = (range,sheet,extra={}) => ({
  id:sheet.id, name:sheet.name, row:range.rowIndex, column:range.columnIndex,
  values:range.values, formulas:range.formulas, r1c1:range.formulasR1C1,
  formats:range.numberFormat, types:range.valueTypes, display:range.text, ...extra,
});
const sourceState = s => [s.row,s.column,s.values,s.formulas,s.formats,s.types,s.display];
const content = s => [s.formulas.map((row,r) => row.map((value,c) =>
  typeof value === 'string' && value.startsWith('=') && value !== s.values[r][c] ? value : s.values[r][c]))];
const checkCancel = signal => { if (signal?.aborted) throw new DOMException('Scan cancelled.', 'AbortError'); };

export async function scan(signal,onProgress,sheetIds,explicit=false) {
  const sheets = await Excel.run(async ctx => {
    const collection = ctx.workbook.worksheets;
    collection.load('items/id,items/name,items/visibility');
    await ctx.sync();
    return collection.items.map(s=>({id:s.id,name:s.name,visibility:s.visibility}));
  });
  const visible=sheets.filter(s=>s.visibility==='Visible');
  const eligible=explicit&&sheetIds?visible:visible.slice(0,12);
  const targets=sheetIds?eligible.filter(s=>sheetIds.has(s.id)):eligible;
  const snapshots = [];
  for (const {id} of targets) {
    checkCancel(signal);
    const snapshot = await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(id);
      sheet.load('id,name');
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('rowIndex,columnIndex,rowCount,columnCount');
      const names=sheet.names, tables=sheet.tables;
      names?.load('items/name');tables?.load('items/name');
      const charts = sheet.charts;
      charts.load('items/name');
      await ctx.sync();
      if (used.isNullObject) return null;
      const columns = Math.min(30,used.columnCount);
      const rows = Math.min(150,used.rowCount,Math.floor(LIMIT/columns));
      const range = sheet.getRange(a1(used.rowIndex,used.columnIndex,rows,columns));
      range.load(properties);
      await ctx.sync();
      return read(range,sheet,{chartCount:charts.items.length,names:names?.items.map(x=>x.name).slice(0,12)??[],
        tables:tables?.items.map(x=>x.name).slice(0,12)??[],partial:rows !== used.rowCount || columns !== used.columnCount});
    });
    if (snapshot) snapshots.push(snapshot);
    onProgress?.(snapshots.length,targets.length);
  }
  checkCancel(signal);
  return {snapshots, sheets, omitted:Math.max(0,sheets.filter(s=>s.visibility==='Visible').length-12)};
}

export async function watchSheets(onChange) {
  if (!supported('1.7')) throw Error('Workbook monitoring requires ExcelApi 1.7.');
  return Excel.run(async ctx=>{
    const list=ctx.workbook.worksheets;
    list.load('items/id,items/visibility');
    await ctx.sync();
    const registrations=list.items.filter(s=>s.visibility==='Visible').slice(0,12).map(s=>{
      const id=s.id;
      return s.onChanged.add(event=>onChange(id,event.address));
    });
    await ctx.sync();
    return async () => Excel.run(ctx,async current=>{
      registrations.forEach(registration=>registration.remove());
      await current.sync();
    });
  });
}

export async function selection() {
  return Excel.run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    const sheet = range.worksheet;
    sheet.load('id,name,visibility');
    range.load('rowIndex,columnIndex,rowCount,columnCount');
    await ctx.sync();
    if (sheet.visibility !== 'Visible') throw Error('Select a visible worksheet.');
    if (range.rowCount*range.columnCount > LIMIT) throw Error('Select up to '+LIMIT+' cells.');
    range.load(properties);
    await ctx.sync();
    return {...read(range,sheet),address:a1(range.rowIndex,range.columnIndex,range.rowCount,range.columnCount)};
  });
}

async function loadSource(ctx,action) {
  const sheet = ctx.workbook.worksheets.getItem(action.sheetId);
  const range = sheet.getRange(action.address);
  sheet.load('id,name,visibility');
  range.load(properties);
  await ctx.sync();
  if (sheet.visibility !== 'Visible') throw Error('Source worksheet is hidden.');
  return {sheet,range,snapshot:read(range,sheet)};
}

export async function inspectRange(sheetId,address) {
  rect(address);
  return Excel.run(async ctx => {
    const {sheet,range,snapshot}=await loadSource(ctx,{sheetId,address});
    // Excel's native selection is a view change and leaves cell formatting alone.
    if (typeof sheet.activate !== 'function' || typeof range.select !== 'function')
      return {snapshot,selected:false};
    try {
      sheet.activate();
      range.select();
      await ctx.sync();
      return {snapshot,selected:true};
    } catch {
      return {snapshot,selected:false};
    }
  });
}

function chartPosition(action) {
  if (action.position) return action.position;
  const {r,c,columns} = rect(action.address);
  return action.destination && action.destination !== action.sheetId ? 'B2:J18' :
    a1(Math.min(r,1048559),Math.min(c+columns+1,16375),17,9);
}
async function chartSlot(ctx,sheet,address) {
  const range = sheet.getRange(address);
  range.load('left,top,width,height');
  const used = range.getUsedRangeOrNullObject(true);
  used.load('address');
  sheet.charts.load('items/left,items/top,items/width,items/height');
  await ctx.sync();
  if (!used.isNullObject) throw Error('The chart position covers occupied cells. Choose an empty rectangle.');
  const overlap = sheet.charts.items.some(c => c.left < range.left+range.width && c.left+c.width > range.left &&
    c.top < range.top+range.height && c.top+c.height > range.top);
  if (overlap) throw Error('The chart position overlaps an existing chart.');
}

function checkPriorFormula(snapshot,action) {
  const prior = snapshot.values[0][0];
  if (snapshot.types[0][0] === 'Error' ||
      (formula(snapshot,0,0) ? !correction(snapshot.r1c1[0][0],rect(action.address).c) :
        prior !== '' && prior !== null && typeof prior !== 'number'))
    throw Error('Only blank, numeric, and simple arithmetic cells can be safely restored by the formula editor.');
}

async function formulaContext(ctx,sheet,action) {
  const {r,c}=rect(action.address);
  if (r<2 || r>1048573) throw Error('Insufficient adjacent rows for formula verification.');
  const target = sheet.getRange(action.address);
  const tables = target.getTables(false);
  tables.load('items/name');
  const offsets = [...action.value.matchAll(/RC\[(-?\d+)\]/g)].map(m=>Number(m[1]));
  const inputs = offsets.map(delta=>sheet.getRange(a1(r,c+delta)));
  inputs.forEach(cell=>cell.load('valueTypes,values,formulas'));
  await ctx.sync();
  if (tables.items.length) throw Error('Formula repairs are unavailable in Excel tables because Excel can autofill calculated columns.');
  if (inputs.some(cell=>!['Integer','Double'].includes(cell.valueTypes[0][0]) || cell.formulas[0][0] !== cell.values[0][0]))
    throw Error('This repair requires two literal numeric inputs in the same row.');
  const band=sheet.getRange(a1(r-2,c,5,1));
  band.load(properties);
  await ctx.sync();
  const snapshot=read(band,sheet);
  if (!correction(action.value,c) || ![0,1,3,4].every(i=>
    formula(snapshot,i,0) && snapshot.r1c1[i][0]===action.value))
    throw Error('The surrounding formula pattern no longer supports this repair.');
  return {band:snapshot,inputs:inputs.map(cell=>cell.values[0][0])};
}

async function checkNonnegative(ctx,sheet,range,snapshot) {
  if (!supported('1.13')) throw Error('This validation requires ExcelApi 1.13 for protected, merged and spill checks.');
  await writable(ctx,sheet,range,true);
  range.dataValidation.load('type');
  await ctx.sync();
  if (range.dataValidation.type!=='None') throw Error('The selection already has a validation rule. Choose cells without validation.');
  if (!snapshot.values.some(([v])=>typeof v==='number' && Number.isFinite(v)) ||
      snapshot.values.some(([v],i)=>snapshot.types[i][0]==='Error'|| (v!=='' && v!==null && (typeof v!=='number'||v<0))))
    throw Error('Select numeric cells without headers or existing negative values.');
}

export async function preview(action) {
  validate(action);
  return Excel.run(async ctx => {
    const {snapshot,sheet,range} = await loadSource(ctx,action);
    const visual = ['chart','dashboard'].includes(action.kind);
    if (visual && !supported('1.10')) throw Error('Charts require ExcelApi 1.10.');
    let destination, position, formulaEvidence, currentStyle, report;
    if (action.kind==='nonnegative') await checkNonnegative(ctx,sheet,range,snapshot);
    if (action.kind==='pivot') {
      if (!supported('1.8')) throw Error('Native PivotTables require ExcelApi 1.8.');
      const protection=ctx.workbook.protection;
      protection.load('protected');
      await ctx.sync();
      if (protection.protected) throw Error('Workbook structure is protected.');
      await assertVisible(ctx,range);
      report=pivotPreview(snapshot,action);
    }
    if (['style','number','trim'].includes(action.kind)) await writable(ctx,sheet,range,true);
    if (action.kind==='trim') await checkTrim(ctx,range,snapshot);
    if (action.kind==='style') {
      loadStyle(range);
      await ctx.sync();
      currentStyle=styleState(range);
    }
    if (['negative','heatmap','dataBar'].includes(action.kind)) {
      await writable(ctx,sheet,range,false);
      await assertVisible(ctx,range);
      if (!snapshot.values.every(([v])=>typeof v==='number' && Number.isFinite(v)))
        throw Error('Select numeric values only.');
    }
    if (action.kind === 'formula') {
      checkPriorFormula(snapshot,action);
      await writable(ctx,sheet,sheet.getRange(action.address),true);
      formulaEvidence = await formulaContext(ctx,sheet,action);
    }
    if (action.kind === 'chart') {
      const target = action.destination ? ctx.workbook.worksheets.getItem(action.destination) : sheet;
      target.load('name,visibility');
      await ctx.sync();
      if (target.visibility !== 'Visible') throw Error('Chart destination is hidden.');
      destination = target.name;
      position = chartPosition(action);
      await chartSlot(ctx,target,position);
    }
    const data = visual ? chartData(snapshot, action.type) : null;
    if (visual) await assertVisible(ctx,range);
    if (action.kind==='chart') {
      const target=action.destination?ctx.workbook.worksheets.getItem(action.destination):sheet;
      await writable(ctx,target,null,false);
    }
    if (action.kind === 'dashboard') {
      if (!additiveMeasure(data.y,action)) throw Error('Dashboard totals require an additive metric.');
      const protection=ctx.workbook.protection;
      protection.load('protected');
      await ctx.sync();
      if (protection.protected) throw Error('Workbook structure is protected.');
    }
    return {action:structuredClone(action), snapshot, data, destination, position, formulaEvidence,currentStyle,report};
  });
}

async function writable(ctx,sheet,range,cellOperation) {
  sheet.protection.load('protected');
  if (!cellOperation) {
    await ctx.sync();
    if (sheet.protection.protected) throw Error('Worksheet is protected.');
    return;
  }
  if (!supported('1.13')) throw Error('Cell edits require ExcelApi 1.13 for merged-cell and spill checks.');
  const merged = range.getMergedAreasOrNullObject();
  merged.load('address');
  range.load('hasSpill,rowHidden,columnHidden');
  await ctx.sync();
  if (sheet.protection.protected || !merged.isNullObject || range.hasSpill !== false ||
      range.rowHidden !== false || range.columnHidden !== false)
    throw Error('Choose a visible, unmerged, non-spilling range on an unprotected worksheet.');
}

async function assertVisible(ctx,range) {
  range.load('rowHidden,columnHidden');
  await ctx.sync();
  if (range.rowHidden !== false || range.columnHidden !== false) throw Error('Chart sources must be visible and unfiltered.');
}

async function checkTrim(ctx,range,snapshot) {
  const tables=range.getTables(false);
  tables.load('items/name');
  await ctx.sync();
  const before=snapshot.values[0][0];
  const after=trimText(before);
  if (tables.items.length || snapshot.types[0][0]!=='String' || formula(snapshot,0,0) || after===null)
    throw Error('Choose eligible literal text outside an Excel table.');
  return after;
}

function styleState(range) {
  return [range.format.font.bold,range.format.font.color,range.format.fill.color];
}
function loadStyle(range) {
  range.format.font.load('bold,color');
  range.format.fill.load('color');
}
function setStyle(range,[bold,font,fill]) {
  range.format.font.bold = bold;
  range.format.font.color = font;
  if (fill) range.format.fill.color = fill;
  else range.format.fill.clear();
}

async function chart(ctx, target, source, action, position) {
  // Seed Excel with the measure column only so there is exactly one data series.
  const measure = source.getColumn(1).getResizedRange(-1,0).getOffsetRange(1,0);
  const keys = source.getColumn(0).getResizedRange(-1,0).getOffsetRange(1,0);
  const item = target.charts.add(chartTypes[action.type],measure,'Columns');
  item.load('name,chartType,left,top,width,height');
  item.series.load('items/name');
  await ctx.sync();
  if (item.series.items.length !== 1) throw Error('Excel produced an unexpected number of chart series.');
  const series = item.series.items[0];
  series.setXAxisValues(keys);
  series.setValues(measure);
  series.name = action.yLabel || 'Measure';
  if (action.color && !['pie','doughnut'].includes(action.type)) {
    if (action.type === 'line') series.format.line.color = action.color;
    else if (action.type === 'scatter') {
      series.markerForegroundColor = action.color;
      series.markerBackgroundColor = action.color;
    }
    else series.format.fill.setSolidColor(action.color);
  }
  item.title.text = action.title;
  item.legend.visible = action.legend;
  if (!['pie','doughnut'].includes(action.type)) {
    item.axes.categoryAxis.title.text = action.xLabel || '';
    item.axes.valueAxis.title.text = action.yLabel || '';
  }
  const [start,end] = position.split(':');
  item.setPosition(target.getRange(start),target.getRange(end));
  item.title.load('text');
  item.legend.load('visible');
  item.load('name,chartType,left,top,width,height');
  series.load('name');
  const coordinate = target.getRange(start);
  coordinate.load('left,top');
  const axes = !['pie','doughnut'].includes(action.type);
  if (axes) {
    item.axes.categoryAxis.title.load('text');
    item.axes.valueAxis.title.load('text');
  }
  await ctx.sync();
  if (!item.name || item.chartType !== chartTypes[action.type] || item.title.text !== action.title ||
      item.legend.visible !== action.legend || !(item.width>0 && item.height>0) || !Number.isFinite(coordinate.left) || !Number.isFinite(coordinate.top) ||
      Math.abs(item.left-coordinate.left)>2 || Math.abs(item.top-coordinate.top)>2 ||
      series.name !== (action.yLabel || 'Measure') ||
      (axes && (item.axes.categoryAxis.title.text !== (action.xLabel || '') ||
                item.axes.valueAxis.title.text !== (action.yLabel || ''))))
    throw Error('Excel chart settings did not verify. Inspect the destination for a partially created chart.');
  return item.name;
}

async function dashboard(ctx,source,action,data,sourceName) {
  const wb = ctx.workbook;
  wb.protection.load('protected');
  const sheets = wb.worksheets;
  sheets.load('items/name');
  await ctx.sync();
  if (wb.protection.protected) throw Error('Workbook structure is protected.');
  const names = new Set(sheets.items.map(s=>s.name.toLowerCase()));
  let name = 'SheetGuard Overview';
  for (let n = 2; names.has(name.toLowerCase()); n++) name = 'SheetGuard Overview '+n;
  const sheet = sheets.add(name);
  const quoted = "'" + sourceName.replaceAll("'","''") + "'!";
  const {r,c,rows} = rect(action.address);
  const measure = quoted + a1(r+1,c+1,rows-1);
  const local = sheet.getRange(a1(5,0,rows,2));
  const links = Array.from({length:rows},(_,i)=>[0,1].map(j=>'='+quoted+a1(r+i,c+j)));
  const kpis = sheet.getRange('B3:B4');
  sheet.getRange('A1').values = [[action.title]];
  sheet.getRange('A3:A4').values = [['Total'],['Observations']];
  kpis.formulas = [['=SUM('+measure+')'],['=COUNT('+measure+')']];
  local.formulas = links;
  local.numberFormat = source.numberFormat;
  sheet.getRange('A1:B1').format.font.bold = true;
  sheet.getRange('A1:B1').format.fill.color = '#d9eee3';
  local.format.autofitColumns();
  kpis.load('values,valueTypes');
  local.load('values,valueTypes');
  await ctx.sync();
  const expected = data.points.reduce((sum,p)=>sum+p[1],0);
  if (Math.abs(kpis.values[0][0]-expected)>1e-9*Math.max(1,Math.abs(expected)) ||
      kpis.values[1][0] !== data.points.length || kpis.valueTypes.some(row=>row.includes('Error')) ||
      !equal(local.values,source.values)) throw Error('Dashboard links or KPIs did not verify.');
  const chartName = await chart(ctx,sheet,local,{...action,xLabel:data.x,yLabel:data.y},'D3:L20');
  return name + ' · ' + chartName;
}

async function createPivot(ctx,range,snapshot,action) {
  if (!supported('1.8')) throw Error('Native PivotTables require ExcelApi 1.8.');
  await assertVisible(ctx,range);
  const report=pivotPreview(snapshot,action), wb=ctx.workbook;
  const sheets=wb.worksheets;
  sheets.load('items/name');wb.protection.load('protected');
  wb.pivotTables.load('items/name');
  await ctx.sync();
  if (wb.protection.protected) throw Error('Workbook structure is protected.');
  const names=new Set([...sheets.items,...wb.pivotTables.items].map(x=>x.name.toLowerCase()));
  let suffix=1;
  while (names.has(('Forma report '+suffix).toLowerCase()) || names.has(('FormaReport'+suffix).toLowerCase())) suffix++;
  const sheetName='Forma report '+suffix, pivotName='FormaReport'+suffix;
  const target=sheets.add(sheetName);
  const pivot=target.pivotTables.add(pivotName,range,target.getRange('A3'));
  pivot.rowHierarchies.add(pivot.hierarchies.getItem(report.dimension));
  const measure=pivot.dataHierarchies.add(pivot.hierarchies.getItem(report.measure));
  measure.summarizeBy='Sum';
  pivot.load('name');
  pivot.rowHierarchies.load('items/name');
  pivot.dataHierarchies.load('items/summarizeBy');
  const body=pivot.layout.getDataBodyRange();
  body.load('values');
  await ctx.sync();
  const got=body.values.flat();
  const want=report.groups.map(g=>g.sum);
  const near=(a,b)=>typeof a==='number' && Math.abs(a-b)<=1e-9*Math.max(1,Math.abs(b));
  const ordered=[...want].sort((a,b)=>a-b);
  const matches=rows=>rows.length===want.length && [...rows].sort((a,b)=>a-b)
    .every((v,i)=>near(v,ordered[i]));
  const totalsMatch=matches(got) || (got.length===want.length+1 &&
    near(got.at(-1),want.reduce((a,b)=>a+b,0)) && matches(got.slice(0,-1)));
  if (pivot.name!==pivotName || pivot.rowHierarchies.items[0]?.name!==report.dimension ||
      pivot.dataHierarchies.items[0]?.summarizeBy!=='Sum' || !totalsMatch)
    throw Error('The PivotTable was created but its output could not be verified. Inspect '+sheetName+' before retrying.');
  return {description:sheetName+' · '+pivotName};
}

export async function apply(p) {
  const {action} = p;
  validate(action);
  return Excel.run(async ctx => {
    const {sheet,range,snapshot} = await loadSource(ctx,action);
    if (!equal(sourceState(snapshot),sourceState(p.snapshot))) throw Error('Source changed. Preview the operation again.');
    if (action.kind==='pivot') return createPivot(ctx,range,snapshot,action);
    if (action.kind==='nonnegative') {
      await checkNonnegative(ctx,sheet,range,snapshot);
      const validation=range.dataValidation;
      validation.rule={decimal:{operator:'GreaterThanOrEqualTo',formula1:0}};
      validation.ignoreBlanks=true;
      validation.errorAlert={showAlert:true,style:'Stop',title:'Enter a non-negative number',message:'Enter zero or a positive number.'};
      validation.load('type,rule,ignoreBlanks,errorAlert');range.load(properties);
      await ctx.sync();
      const rule=validation.rule?.decimal, threshold=rule?.formula1;
      if (validation.type!=='Decimal'||rule?.operator!=='GreaterThanOrEqualTo'||
          ![0,'0','=0'].includes(threshold)||!validation.ignoreBlanks||validation.errorAlert?.style!=='Stop'||
          !validation.errorAlert?.showAlert||!equal(sourceState(read(range,sheet)),sourceState(snapshot)))
        throw Error('Excel did not verify the validation rule. Inspect the range before retrying.');
      return {description:sheet.name+'!'+action.address};
    }
    if (['chart','dashboard'].includes(action.kind)) {
      if (!supported('1.10')) throw Error('Charts require ExcelApi 1.10.');
      const data = chartData(snapshot,action.type);
      await assertVisible(ctx,range);
      if (action.kind === 'dashboard') {
          if (!additiveMeasure(data.y,action))
          throw Error('Dashboard totals require an explicitly additive heading.');
        return {description:await dashboard(ctx,range,action,data,snapshot.name)};
      }
      const target = action.destination ? ctx.workbook.worksheets.getItem(action.destination) : sheet;
      target.load('name,visibility');
      await writable(ctx,target,null,false);
      if (target.visibility !== 'Visible' || target.name !== p.destination || chartPosition(action) !== p.position)
        throw Error('Chart destination changed. Preview again.');
      await chartSlot(ctx,target,p.position);
      const name = await chart(ctx,target,range,{...action,xLabel:action.xLabel ?? data.x,yLabel:action.yLabel ?? data.y},p.position);
      range.load(properties);
      await ctx.sync();
      if (!equal(content(snapshot),content(read(range,sheet)))) throw Error('Chart created, but the source contents changed during verification.');
      return {description:target.name+' · '+name};
    }
    if (['heatmap','negative','dataBar'].includes(action.kind)) {
      if (!supported('1.6')) throw Error('Conditional formatting requires ExcelApi 1.6.');
      await writable(ctx,sheet,range,false);
      await assertVisible(ctx,range);
      if (!snapshot.values.every(([v])=>typeof v === 'number' && Number.isFinite(v))) throw Error('Select numeric values only.');
      const type = {heatmap:'ColorScale',negative:'CellValue',dataBar:'DataBar'}[action.kind];
      const rule = range.conditionalFormats.add(type);
      if (action.kind === 'negative') {
        rule.cellValue.rule = {operator:'LessThan',formula1:'=0'};
        rule.cellValue.format.fill.color = '#ffd9df';
      } else if (action.kind === 'heatmap') {
        rule.colorScale.criteria = {
          minimum:{formula:null,type:'LowestValue',color:'#e3eef7'},
          maximum:{formula:null,type:'HighestValue',color:'#429b76'},
        };
      }
      rule.load('id,type');
      range.load(properties);
      await ctx.sync();
      if (!rule.id || rule.type !== type || !equal(content(snapshot),content(read(range,sheet))))
        throw Error('Conditional formatting could not be verified.');
      return {description:'Conditional format '+rule.id};
    }

    await writable(ctx,sheet,range,true);
    if (action.kind==='formula' && !equal(await formulaContext(ctx,sheet,action),p.formulaEvidence))
      throw Error('The formula context changed. Preview again.');
    let nextStyle, trimmed;
    if (action.kind === 'style') {
      loadStyle(range);
      await ctx.sync();
      if (!equal(styleState(range),p.currentStyle)) throw Error('The source style changed. Preview the change again.');
      nextStyle = headingStyle(action);
      setStyle(range,nextStyle);
    } else if (action.kind==='trim') {
      trimmed=await checkTrim(ctx,range,snapshot);
      range.values=[[trimmed]];
    } else if (action.kind === 'number') {
      range.numberFormat = snapshot.values.map(row=>row.map(()=>action.value));
    } else {
      checkPriorFormula(snapshot,action);
      range.formulasR1C1 = [[action.value]];
    }
    range.load(properties);
    if (action.kind === 'style') loadStyle(range);
    await ctx.sync();
    const after = read(range,sheet);
    const verified = action.kind === 'formula' ? after.r1c1[0][0] === action.value && after.types[0][0] !== 'Error'
      : action.kind==='trim'?after.values[0][0]===trimmed && after.types[0][0]==='String' &&
        !formula(after,0,0) && equal(after.formats,snapshot.formats)
      : equal(content(snapshot),content(after)) && (action.kind === 'number'
        ? after.formats.every(row=>row.every(value=>value===action.value))
        : equal(styleState(range).map(v=>typeof v==='string'?v.toUpperCase():v),nextStyle));
    if (!verified) {
      // A known formula write can be reverted in the same operation if it produced an error.
      if (action.kind === 'formula' && after.r1c1[0][0] === action.value) {
        range.formulasR1C1 = snapshot.r1c1;
        range.load(properties);
        await ctx.sync();
        if (equal(content(read(range,sheet)),content(snapshot))) throw Error('The new formula produced an error. The original cell was restored.');
      }
      throw Error('Workbook write did not verify. Inspect the target before retrying.');
    }
    return {description:sheet.name+'!'+action.address,
      recovery:action.kind === 'style' ? undefined : {action,before:snapshot,after}};
  });
}

export async function rollback(recovery) {
  const {action,before,after} = recovery;
  return Excel.run(async ctx => {
    const {sheet,range,snapshot} = await loadSource(ctx,action);
    await writable(ctx,sheet,range,true);
    if (!equal(sourceState(snapshot),sourceState(after))) throw Error('The range changed after applying the operation. Undo was skipped.');
    if (action.kind === 'number') range.numberFormat = before.formats;
    else if (action.kind==='trim') {
      const tables=range.getTables(false);
      tables.load('items/name');
      await ctx.sync();
      if (tables.items.length || trimText(before.values[0][0])!==snapshot.values[0][0])
        throw Error('This text cell cannot be safely restored.');
      range.values=before.values;
    }
    else if (action.kind === 'formula') {
      await formulaContext(ctx,sheet,action);
      range.formulasR1C1 = before.r1c1;
    } else throw Error('This operation cannot be undone from SheetGuard.');
    range.load(properties);
    await ctx.sync();
    const restored = read(range,sheet);
    if (!equal(content(restored),content(before)) || !equal(restored.formats,before.formats)) throw Error('Undo could not be verified.');
  });
}

// Long selections are loaded in bounded blocks. No cells are changed.
export async function readDataset(signal,onProgress,target,expand=false) {
  return Excel.run(async ctx=>{
    if (target) rect(target.address,DATASET_CELL_LIMIT);
    let range=target?ctx.workbook.worksheets.getItem(target.sheetId).getRange(target.address):ctx.workbook.getSelectedRange();
    const sheet=range.worksheet;
    const dimensions='rowIndex,columnIndex,rowCount,columnCount';
    range.load(dimensions);sheet.load('id,name,visibility');
    await ctx.sync();
    if (expand && range.rowCount===1 && range.columnCount===1 && supported('1.7') &&
        typeof range.getSurroundingRegion==='function') {
      range=range.getSurroundingRegion();range.load(dimensions);await ctx.sync();
    }
    const {rowIndex:r,columnIndex:c,rowCount:rows,columnCount:columns}=range;
    if (sheet.visibility!=='Visible') throw Error('Select a range on a visible worksheet.');
    if (rows<2 || columns<1 || rows>DATASET_ROW_LIMIT || rows*columns>DATASET_CELL_LIMIT)
      throw Error('Select a header and data within 50,000 rows and 160,000 cells.');
    const values=[],errors=[];
    const blockSize=Math.max(1,Math.min(2000,Math.floor(8000/columns)));
    let textSize=0;
    for (let offset=0;offset<rows;offset+=blockSize) {
      checkCancel(signal);
      const block=sheet.getRange(a1(r+offset,c,Math.min(blockSize,rows-offset),columns));
      block.load('values,valueTypes');await ctx.sync();
      for (let i=0;i<block.values.length;i++) {
        const row=block.values[i];
        textSize+=row.reduce((n,v)=>n+(typeof v==='string'?v.length:0),0);
        if (textSize>4000000) throw Error('This selection contains too much text. Choose fewer columns or rows.');
        values.push(row);
        errors.push(block.valueTypes[i].flatMap((v,j)=>v==='Error'?[j]:[]));
      }
      onProgress?.(values.length,rows);
    }
    checkCancel(signal);
    return {sheetId:sheet.id,sheet:sheet.name,address:a1(r,c,rows,columns),values,errors};
  });
}


export async function readContext(target) {
  rect(target.address);
  return Excel.run(async ctx=>(await loadSource(ctx,target)).snapshot);
}


export async function selectionAnchor() {
  return Excel.run(async ctx=>{
    const range=ctx.workbook.getSelectedRange(),sheet=range.worksheet;
    range.load('rowIndex,columnIndex,rowCount,columnCount');sheet.load('id,name,visibility');
    await ctx.sync();
    if (sheet.visibility!=='Visible') throw Error('Select cells on a visible worksheet.');
    return {id:sheet.id,name:sheet.name,address:a1(range.rowIndex,range.columnIndex,range.rowCount,range.columnCount),
      rows:range.rowCount,columns:range.columnCount};
  });
}

export async function activeWorksheet() {
  return Excel.run(async ctx=>{
    const sheet=ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('id,name');
    await ctx.sync();
    return {id:sheet.id,name:sheet.name};
  });
}
