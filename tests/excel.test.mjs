import test from 'node:test';
import assert from 'node:assert/strict';
import {rect,a1} from '../public/engine.mjs';
import {preview,apply,rollback} from '../public/excel.mjs';

function host({protectedSheet=false,table=false,merged=false,validationType='None'}={}) {
  const values = Array.from({length:6},(_,r)=>['Item '+r,2,3,6]);
  const formulas = values.map((row,r)=>[...row.slice(0,3),'=B'+(r+1)+'*C'+(r+1)]);
  const r1c1 = values.map(row=>[...row.slice(0,3),'=RC[-2]*RC[-1]']);
  values[3][3]=5;formulas[3][3]='=B4+C4';r1c1[3][3]='=RC[-2]+RC[-1]';
  let writes=0;
  const validation={type:validationType,load(){},set rule(v){writes++;this.type='Decimal';this.savedRule=v;},
    get rule(){return this.savedRule;}};
  const matrix = (rows,r,c,n,m)=>rows.slice(r,r+n).map(row=>row.slice(c,c+m));
  const sheet={
    id:'s',name:'Data',visibility:'Visible',load(){},
    protection:{protected:protectedSheet,load(){}},
    getRange(address) {
      const {r,c,rows,columns}=rect(address);
      const read=items=>matrix(items,r,c,rows,columns);
      return {load(){},rowIndex:r,columnIndex:c,rowHidden:false,columnHidden:false,hasSpill:false,
        getMergedAreasOrNullObject:()=>({load(){},isNullObject:!merged}),
        dataValidation:validation,getTables:()=>({load(){},items:table?[{}]:[]}),
        get values(){return read(values);},
        set values(grid){writes++;values[r][c]=grid[0][0];formulas[r][c]=grid[0][0];r1c1[r][c]=grid[0][0];},
        get formulas(){return read(formulas);},
        get formulasR1C1(){return read(r1c1);},
        set formulasR1C1(grid){
          writes++;
          r1c1[r][c]=grid[0][0];
          const op=grid[0][0].includes('+')?'+':'*';
          formulas[r][c]='='+a1(r,1)+op+a1(r,2);
          values[r][c]=op==='+'?values[r][1]+values[r][2]:values[r][1]*values[r][2];
        },
        get text(){return read(values).map(row=>row.map(String));},
        get numberFormat(){return read(values).map(row=>row.map(()=> 'General'));},
        get valueTypes(){return read(values).map(row=>row.map(v=>typeof v==='number'?'Double':'String'));},
      };
    },
  };
  globalThis.Office={context:{requirements:{isSetSupported:()=>true}}};
  globalThis.Excel={run:func=>func({sync:async()=>{},workbook:{worksheets:{getItem:()=>sheet}}})};
  return {values,formulas,r1c1,writes:()=>writes};
}
const action={kind:'formula',sheetId:'s',address:'D4',value:'=RC[-2]*RC[-1]'};

test('formula preview never writes, acceptance verifies the repair, and undo restores the previous formula', async()=>{
  const workbook=host();
  const proposal=await preview(action);
  assert.equal(workbook.writes(),0);
  const result=await apply(proposal);
  assert.equal(workbook.formulas[3][3],'=B4*C4');
  assert.equal(workbook.values[3][3],6);
  await rollback(result.recovery);
  assert.equal(workbook.formulas[3][3],'=B4+C4');
  assert.equal(workbook.writes(),2);
});

test('formula approvals cannot proceed with stale inputs or a changed adjacent pattern', async()=>{
  const workbook=host();
  const proposal=await preview(action);
  workbook.values[3][1]=4;workbook.formulas[3][1]=4;
  await assert.rejects(apply(proposal),/context changed/);
  workbook.r1c1[2][3]='=RC[-2]+RC[-1]';
  await assert.rejects(preview(action),/no longer supports/);
  assert.equal(workbook.writes(),0);
});

test('protected, merged and table formula targets are rejected before a write', async()=>{
  for (const flags of [{protectedSheet:true},{merged:true},{table:true}]) {
    const workbook=host(flags);
    await assert.rejects(preview(action),/unprotected worksheet|Excel tables/);
    assert.equal(workbook.writes(),0);
  }
});

test('undo refuses to overwrite a later change', async()=>{
  const workbook=host();
  const result=await apply(await preview(action));
  workbook.values[3][3]=17;workbook.formulas[3][3]=17;
  await assert.rejects(rollback(result.recovery),/changed after applying/);
  assert.equal(workbook.values[3][3],17);
  assert.equal(workbook.writes(),1);
});

test('visual formula preview shows a translated A1 formula and evaluated numeric estimate',async()=>{
  const {formulaPreview}=await import('../public/engine.mjs');
  const workbook=host();
  const item=formulaPreview(await preview(action));
  assert.equal(item.before,'=B4+C4');
  assert.equal(item.after,'=B4*C4');
  assert.equal(item.current,'5');
  assert.equal(item.estimate,6);
  assert.deepEqual(item.neighbours.map(([cell])=>cell),['D2','D3','D5','D6']);
  assert.equal(workbook.writes(),0);
});


test('literal trim is preview-only until approval, preserves format, and supports undo',async()=>{
  const {previewRows}=await import('../public/engine.mjs');
  const workbook=host();
  workbook.values[3][0]=workbook.formulas[3][0]=workbook.r1c1[3][0]='  Audio  ';
  const proposed=await preview({kind:'trim',sheetId:'s',address:'A4'});
  assert.equal(previewRows(proposed)[0].after,'Audio');
  assert.equal(workbook.writes(),0);
  const applied=await apply(proposed);
  assert.equal(workbook.values[3][0],'Audio');
  await rollback(applied.recovery);
  assert.equal(workbook.values[3][0],'  Audio  ');
  assert.equal(workbook.writes(),2);
});

test('large-sheet scans bound reads and incremental scans visit only the changed visible sheet',async()=>{
  const {scan}=await import('../public/excel.mjs');
  const requests=[];
  const grid=(rows,cols,value)=>Array.from({length:rows},()=>Array(cols).fill(value));
  function worksheet(id,visibility) {
    return {id,name:id,visibility,load(){},charts:{items:[],load(){}},
      getUsedRangeOrNullObject(){return {rowIndex:0,columnIndex:0,rowCount:350000,columnCount:200,load(){}};},
      getRange(address){
        requests.push([id,address]);const {r,c,rows,columns}=rect(address);const blank=grid(rows,columns,'');
        return {rowIndex:r,columnIndex:c,values:blank,formulas:blank,formulasR1C1:blank,
          numberFormat:blank,valueTypes:blank,text:blank,load(){}};
      },
    };
  }
  const list=[worksheet('First','Visible'),worksheet('Second','Visible'),worksheet('Hidden','Hidden')];
  globalThis.Excel={run:fn=>fn({sync:async()=>{},workbook:{
    worksheets:{load(){},items:list,getItem:id=>list.find(s=>s.id===id)},
  }})};
  const result=await scan(undefined,undefined,new Set(['Second','Hidden']));
  assert.deepEqual(requests,[['Second','A1:AD80']]);
  assert.equal(result.snapshots[0].partial,true);
  assert.equal(result.sheets.find(s=>s.id==='Hidden').visibility,'Hidden');
});

test('monitoring subscribes only visible sheets and removes the handlers on pause',async()=>{
  const {watchSheets}=await import('../public/excel.mjs');
  host();
  const registered=[], removed=[];
  const worksheets=['Visible','Hidden'].map(visibility=>({id:visibility,visibility,onChanged:{
    add(handler){registered.push(handler);return {remove(){removed.push(visibility);}};},
  }}));
  const ctx={sync:async()=>{},workbook:{worksheets:{load(){},items:worksheets}}};
  globalThis.Excel={run:(first,second)=>(second??first)(ctx)};
  const changes=[];
  const stop=await watchSheets((...args)=>changes.push(args));
  assert.equal(registered.length,1);
  registered[0]({address:'B7'});
  assert.deepEqual(changes,[['Visible','B7']]);
  await stop();
  assert.deepEqual(removed,['Visible']);
});

test('selected dataset is read in bounded chunks and can be cancelled between blocks',async()=>{
  const {readDataset}=await import('../public/excel.mjs');
  const reads=[];
  const sheet={id:'s',name:'Data',visibility:'Visible',load(){},
    getRange(address) {
      reads.push(address);
      const [,a,b]=/^A(\d+):B(\d+)$/.exec(address);
      const values=Array.from({length:Number(b)-Number(a)+1},()=>['East',4]);
      return {values,valueTypes:values.map(()=>['String','Double']),load(){}};
    }};
  const range={worksheet:sheet,rowIndex:0,columnIndex:0,rowCount:12000,columnCount:2,load(){}};
  globalThis.Excel={run:fn=>fn({sync:async()=>{},workbook:{getSelectedRange:()=>range}})};
  const cut=new AbortController();
  await assert.rejects(readDataset(cut.signal,()=>cut.abort()),{name:'AbortError'});
  assert.deepEqual(reads,['A1:B2000']);
  reads.length=0;
  const data=await readDataset();
  assert.equal(data.values.length,12000);assert.equal(reads.length,6);
});


test('validation preview performs no write; apply verifies a native rule and rejects existing validation',async()=>{
  const workbook=host();
  const action={kind:'nonnegative',sheetId:'s',address:'B2:B5'};
  const plan=await preview(action);
  assert.equal(workbook.writes(),0);
  await apply(plan);
  assert.equal(workbook.writes(),1);
  await assert.rejects(preview(action),/already has a validation rule/);
  assert.equal(workbook.values[1][1],2);
});
