import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
import {FileBlob,SpreadsheetFile,Workbook} from '@oai/artifact-tool';
import {conditionalQuery} from '../public/query-plan.mjs';
import {rankedAnalysis} from '../public/agent.mjs';

const [output,previews]=process.argv.slice(2);
if (!output) throw Error('Usage: node tools/build-live-fixtures.mjs OUTPUT_DIR [PREVIEW_DIR]');

const packs=[
  {name:'Forma verification fixtures',cases:[
    {name:'Work orders',group:'Workshop',measure:'Hours',flag:'State',period:'Priority',r:3,c:1,
      labels:['Forge','Assembly','Packing'],active:'Open',other:'Closed',table:true},
    {name:'Electives',group:'Department',measure:'Credits',flag:'Enrolled',period:'Year',r:1,c:0,
      labels:['Science','Music','History'],active:true,other:false,formula:true},
    {name:'Warehouse',group:'Aisle',measure:'Units',flag:'Released',period:'Cycle',r:5,c:2,
      labels:['Loft','East','West'],active:true,other:false,textNumber:true},
    {name:'Survey runs',group:'Station',measure:'Samples',flag:'Reviewed',period:'Batch',r:0,c:3,
      labels:['Pond','Slope','Meadow'],active:'Yes',other:'No'},
  ]},
  {name:'Forma holdout fixtures',cases:[
    {name:'Dispatch',group:'Route',measure:'Packages',flag:'Loaded',period:'Wave',r:1,c:2,
      labels:['Hill','Valley','Coast'],active:true,other:false,table:true},
    {name:'Archive Q2',group:'Collection',measure:'Pages',flag:'Digitised',period:'Shift',r:4,c:1,
      labels:['Letters','Ledgers','Maps'],active:'Yes',other:'No',formula:true},
    {name:"Ops '東",group:'班',measure:'件数',flag:'確認',period:'段階',r:2,c:3,
      labels:['東','西','北'],active:true,other:false,textNumber:true},
    {name:'Lab allocation',group:'Bench',measure:'Minutes',flag:'Status',period:'Run',r:0,c:0,
      labels:['Quartz','Slate','Basalt'],active:'Ready',other:'Waiting'},
  ]},
];

const column=index=>String.fromCharCode(65+index);
const addresses=(firstRow,firstCol,rows,columns)=>column(firstCol)+(firstRow+1)+':'+column(firstCol+columns-1)+(firstRow+rows);
const expected=[];
await mkdir(output,{recursive:true});
if (previews) await mkdir(previews,{recursive:true});
for (const [packIndex,pack] of packs.entries()) {
  const book=Workbook.create();
  const entries=[];
  for (const [caseIndex,f] of pack.cases.entries()) {
    const sheet=book.worksheets.add(f.name);
    sheet.showGridLines=false;
    const roles=(packIndex+caseIndex)%2 ? ['flag','group','period','measure'] : ['group','measure','flag','period'];
    const delta=1+packIndex*17+caseIndex*7;
    const records=[
      {group:f.labels[0],measure:19+delta,flag:f.active,period:2},
      {group:f.labels[1],measure:81+delta,flag:f.other,period:3},
      {group:f.labels[0],measure:-3,flag:f.active,period:4},
      {group:f.labels[2],measure:20+delta,flag:f.active,period:1},
      {group:f.labels[1],measure:10+delta,flag:f.active,period:3},
      {group:f.labels[2],measure:4+delta,flag:f.active,period:2},
      {group:f.labels[0],measure:f.textNumber?'45':0,flag:f.active,period:3},
      {group:f.labels[1],measure:null,flag:f.active,period:3},
    ];
    const matrix=[roles.map(key=>f[key]),...records.map(record=>roles.map(key=>record[key]))];
    const address=addresses(f.r,f.c,matrix.length,roles.length);
    const range=sheet.getRange(address);
    range.values=matrix;
    range.format.font.name='Arial';
    range.format.font.size=11;
    range.format.columnWidthPx=126;
    range.format.rowHeight=18;
    const header=range.getRow(0);
    header.format.fill='#234E3D';
    header.format.font.color='#FFFFFF';
    header.format.font.bold=true;
    if (f.r>=2) {
      const title=sheet.getCell(f.r-2,f.c);
      title.values=[[f.name]];
      title.format.font.bold=true;
    }
    const measureIndex=roles.indexOf('measure');
    if (f.formula) sheet.getCell(f.r+1,f.c+measureIndex).formulas=[['='+delta+'+19']];
    if (f.table) sheet.tables.add(address,true,'InputRecords'+(caseIndex+1));

    // The oracle operates on the generating records before workbook authoring or column permutation.
    const groups=new Map();
    for (const [index,record] of records.entries()) {
      if (record.flag!==f.active||record.period<2||typeof record.measure!=='number') continue;
      const group=groups.get(record.group)??{label:record.group,total:0,measureCells:[],filterCells:[]};
      const row=f.r+index+2;
      group.total+=record.measure;
      group.measureCells.push(column(f.c+measureIndex)+row);
      group.filterCells.push([column(f.c+roles.indexOf('flag'))+row,column(f.c+roles.indexOf('period'))+row]);
      groups.set(record.group,group);
    }
    entries.push({workbook:pack.name+'.xlsx',sheet:f.name,selection:address,
      question:'Rank '+JSON.stringify(f.group)+' by total '+JSON.stringify(f.measure)+' where '+
        JSON.stringify(f.flag)+' = '+JSON.stringify(f.active)+' and '+JSON.stringify(f.period)+' >= 2',
      expected:[...groups.values()].sort((a,b)=>b.total-a.total)});
  }
  book.recalculate();
  for (const [index,entry] of entries.entries()) {
    const f=pack.cases[index],sheet=book.worksheets.getItem(entry.sheet);
    const {values}=sheet.getRange(entry.selection);
    assert.equal(values.length,9);
    const measureIndex=values[0].indexOf(f.measure);
    assert.equal(values[1][measureIndex],20+packIndex*17+index*7);
    assert.equal(typeof values[7][measureIndex],f.textNumber?'string':'number');
    const reads=Object.fromEntries(entry.expected.map(group=>[group.label,
      group.measureCells.reduce((sum,cell)=>sum+sheet.getRange(cell).values[0][0],0)]));
    assert.deepEqual(reads,Object.fromEntries(entry.expected.map(group=>[group.label,group.total])));
    const inspection=await book.inspect({kind:'match',sheetId:entry.sheet,range:entry.selection,
      searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
      options:{useRegex:true,maxResults:5}});
    assert(!/\"kind\":\"match\"/.test(inspection.ndjson));
    if (previews) {
      const range=addresses(Math.max(f.r-2,0),f.c,Math.min(f.r,2)+9,4);
      const image=await book.render({sheetName:entry.sheet,range,scale:1,format:'png'});
      await writeFile(join(previews,packIndex+'-'+index+'.png'),new Uint8Array(await image.arrayBuffer()));
    }
  }
  const file=join(output,pack.name+'.xlsx');
  await (await SpreadsheetFile.exportXlsx(book)).save(file);
  const reopened=await SpreadsheetFile.importXlsx(await FileBlob.load(file));
  for (const entry of entries) {
    const {plan}=conditionalQuery(entry.question);
    const values=reopened.worksheets.getItem(entry.sheet).getRange(entry.selection).values;
    const result=rankedAnalysis({dimension:plan.group,measure:plan.measure,queryPlan:plan},
      {sheetId:entry.sheet,sheet:entry.sheet,address:entry.selection,values});
    assert.deepEqual(result.evidence.groups.map(group=>({label:group.label,total:group.sum,
      measureCells:group.records.map(record=>record.measureCell),
      filterCells:group.records.map(record=>record.filterCells.map(cell=>cell.address))})),entry.expected);
  }
  expected.push(...entries);
}
await writeFile(join(output,'Forma fixture answers.json'),JSON.stringify(expected,null,2)+'\n');
