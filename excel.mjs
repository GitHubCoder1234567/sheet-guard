import {a1, rect, validate, chartData, chartTypes, palettes, correction, formula, LIMIT} from './engine.mjs';

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

export async function scan(signal,onProgress) {
  const sheets = await Excel.run(async ctx => {
    const collection = ctx.workbook.worksheets;
    collection.load('items/id,items/name,items/visibility');
    await ctx.sync();
    return collection.items.filter(s => s.visibility === 'Visible').map(s => s.id);
  });
  const snapshots = [];
  for (const id of sheets.slice(0,12)) {
    checkCancel(signal);
    const snapshot = await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(id);
      sheet.load('id,name');
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('rowIndex,columnIndex,rowCount,columnCount');
      const charts = sheet.charts;
      charts.load('items/name');
      await ctx.sync();
      if (used.isNullObject) return null;
      const columns = Math.min(30,used.columnCount);
      const rows = Math.min(150,used.rowCount,Math.floor(LIMIT/columns));
      const range = sheet.getRange(a1(used.rowIndex,used.columnIndex,rows,columns));
      range.load(properties);
      await ctx.sync();
      return read(range,sheet,{chartCount:charts.items.length, partial:rows !== used.rowCount || columns !== used.columnCount});
    });
    if (snapshot) snapshots.push(snapshot);
    onProgress?.(snapshots.length,sheets.length);
  }
  checkCancel(signal);
  return {snapshots, omitted:Math.max(0,sheets.length-12)};
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
  return [sourceState(snapshot),inputs.map(cell=>cell.values[0][0])];
}

export async function preview(action) {
  validate(action);
  return Excel.run(async ctx => {
    const {snapshot,sheet} = await loadSource(ctx,action);
    const visual = ['chart','dashboard'].includes(action.kind);
    if (visual && !supported('1.10')) throw Error('Charts require ExcelApi 1.10.');
    let destination, position, formulaEvidence;
    if (action.kind === 'formula') {
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
    if (action.kind === 'dashboard' && !/^(?:total )?(?:sales|revenue|units|quantity|orders)(?: \([A-Z]{3}\))?$/i.test(data.y))
      throw Error('Dashboard totals require an explicitly additive heading.');
    return {action:structuredClone(action), snapshot, data, destination, position, formulaEvidence};
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

export async function apply(p) {
  const {action} = p;
  validate(action);
  return Excel.run(async ctx => {
    const {sheet,range,snapshot} = await loadSource(ctx,action);
    if (!equal(sourceState(snapshot),sourceState(p.snapshot))) throw Error('Source changed. Preview the operation again.');
    if (['chart','dashboard'].includes(action.kind)) {
      if (!supported('1.10')) throw Error('Charts require ExcelApi 1.10.');
      const data = chartData(snapshot,action.type);
      await assertVisible(ctx,range);
      if (action.kind === 'dashboard') {
          if (!/^(?:total )?(?:sales|revenue|units|quantity|orders)(?: \([A-Z]{3}\))?$/i.test(data.y))
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
    let nextStyle;
    if (action.kind === 'style') {
      nextStyle = [true,'#FFFFFF',(action.color || palettes[action.value]).toUpperCase()];
      setStyle(range,nextStyle);
    } else if (action.kind === 'number') {
      range.numberFormat = snapshot.values.map(row=>row.map(()=>action.value));
    } else {
      const prior = snapshot.values[0][0];
      if (snapshot.types[0][0] === 'Error' ||
          (formula(snapshot,0,0) ? !correction(snapshot.r1c1[0][0],rect(action.address).c) :
            prior !== '' && prior !== null && typeof prior !== 'number'))
        throw Error('Only blank, numeric, and simple arithmetic cells can be safely restored by the formula editor.');
      range.formulasR1C1 = [[action.value]];
    }
    range.load(properties);
    if (action.kind === 'style') loadStyle(range);
    await ctx.sync();
    const after = read(range,sheet);
    const verified = action.kind === 'formula' ? after.r1c1[0][0] === action.value && after.types[0][0] !== 'Error'
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
