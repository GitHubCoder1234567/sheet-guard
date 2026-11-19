import {a1, rect, previewRows, formulaPreview, headingStyle, profile, audit, categories, chartTypes, palettes, conflict, validate} from './engine.mjs';
import {scan, watchSheets, selection, inspectRange, preview, apply, rollback, supported} from './excel.mjs';
import {plot} from './plot.mjs';

const nav = document.querySelector('#nav'), main = document.querySelector('#main');
const notice = document.querySelector('#notice'), connection = document.querySelector('#connection');
const tabs = ['Chat','Audit','Recommendations','Visualisations','History','Settings'];
const state = {
  tab:'Chat', ready:false, busy:false, findings:[], snapshots:[], dismissed:new Set(),
  previews:new Map(), edits:new Map(), checked:new Set(), history:[], messages:[],
  filter:'All', sort:'Severity', query:'', selected:null, inspected:new Map(), explanation:new Set(),
  previewMode:new Map(), clarified:new Map(), batchReview:false, lastScan:null, sheets:[], monitor:null, consent:false, draft:'', scanned:false, manualKind:'chart', accessPassword:'',
};
function node(tag,value,cls) {
  const el=document.createElement(tag);
  if (value !== undefined) el.textContent=String(value);
  if (cls) el.className=cls;
  return el;
}
function button(label, fn, {primary=false, enabled=true, during=false}={}) {
  const b=node('button',label,primary?'primary':undefined);
  b.type='button'; b.disabled=!enabled || (!during && state.busy);
  b.addEventListener('click',fn);
  return b;
}
function block(...children) {
  const card=node('article',undefined,'card');
  card.append(...children);return card;
}
function field(label,control) {
  const wrap=node('label',label);wrap.append(control);return wrap;
}
function select(options,value) {
  const el=node('select');
  options.forEach(([v,l])=>{const opt=node('option',l);opt.value=v;el.append(opt);});
  el.value=value;
  return el;
}
function setTab(tab) {state.tab=tab;render();}
function render() {
  nav.replaceChildren(...tabs.map(tab=>{
    const b=button(tab,()=>setTab(tab),{during:true});
    if (tab===state.tab) b.setAttribute('aria-current','page');
    return b;
  }));
  main.replaceChildren(node('h2',state.tab));
  if (!state.ready) {
    main.append(block(node('p','Open this task pane in Microsoft Excel. The manifest and setup guide are in the release package.')));
    return;
  }
  ({Chat:chatView,Audit:auditView,Recommendations:()=>recommendations(false),Visualisations:()=>recommendations(true),
    History:historyView,Settings:settingsView}[state.tab])();
}
async function run(work,label) {
  if (state.busy) return;
  state.busy=true;notice.textContent=label;render();
  try {notice.textContent=await work() || '';}
  catch (error) {notice.textContent=error.message || String(error);}
  finally {state.busy=false;render();}
}
const findingAddress = f => f.sheet+'!'+f.address;
function available(a) {
  if (['chart','dashboard'].includes(a.kind)) return supported('1.10');
  if (['negative','heatmap','dataBar'].includes(a.kind)) return supported('1.6');
  return supported('1.13');
}
async function doScan(sheetIds) {
  state.scanController=new AbortController();
  render();
  try {
    const {snapshots,sheets,omitted}=await scan(state.scanController.signal,(done,total)=>{
      notice.textContent='Scanned '+done+' populated worksheet(s), '+total+' visible.';
    },sheetIds);
    state.snapshots=sheetIds?[...state.snapshots.filter(s=>!sheetIds.has(s.id)),...snapshots]:snapshots;
    state.sheets=sheets;state.omitted=omitted;state.scanned=true;state.lastScan=new Date();
    const manual=state.findings.filter(f=>f.manual);
    state.findings=[...audit(state.snapshots),...manual];
    const active=new Set(state.findings.map(f=>f.id));
    for (const cache of [state.edits,state.previewMode,state.explanation]) {
      for (const id of cache.keys()) if (!active.has(id)) cache.delete(id);
    }
    state.previews.clear();state.checked.clear();state.clarified.clear();state.inspected.clear();state.batchReview=false;
    return 'Audit complete.';
  } finally {state.scanController=null;}
}
const pendingSheets=new Set();
let monitorTimer;
function flushMonitoring() {
  if (!state.monitor || !pendingSheets.size) return;
  if (state.busy || state.scanController) {
    monitorTimer=setTimeout(flushMonitoring,800);return;
  }
  const pending=new Set(pendingSheets);pendingSheets.clear();
  run(()=>doScan(pending),'Refreshing changed worksheets…');
}
async function toggleMonitoring() {
  if (state.monitor) {
    await state.monitor();
    state.monitor=null;pendingSheets.clear();clearTimeout(monitorTimer);
    return 'Workbook monitoring paused.';
  }
  state.monitor=await watchSheets((sheetId)=>{
    pendingSheets.add(sheetId);
    clearTimeout(monitorTimer);
    monitorTimer=setTimeout(flushMonitoring,900);
  });
  return 'Workbook monitoring enabled. Changed sheets will be sampled.';
}
function auditView() {
  const controls=node('div',undefined,'actions');
  controls.append(button('Scan visible sheets',()=>run(doScan,'Scanning workbook…'),{primary:true}));
  if (state.scanController) controls.append(button('Cancel scan',()=>state.scanController.abort(),{during:true}));
  controls.append(button(state.monitor?'Pause monitoring':'Enable monitoring',
    ()=>run(toggleMonitoring,'Updating monitoring…'),{enabled:supported('1.7')}));
  main.append(controls);
  if (!state.scanned) {main.append(block(node('p','Scan the workbook for formula patterns, data quality issues and visualisation opportunities.')));return;}
  const stats=node('div',undefined,'metrics');
  const cells=state.snapshots.reduce((n,s)=>n+s.values.length*s.values[0].length,0);
  const metrics=[['Worksheets',state.snapshots.length],['Cells',cells],
    ...['Errors','Formatting','Data Quality','Visualisations','Analysis'].map(c=>[c,state.findings.filter(f=>!f.manual && f.category===c).length])];
  for (const [label,value] of metrics) {
    const item=node('div',undefined,'metric');
    item.append(node('strong',value),node('span',label));stats.append(item);
  }
  main.append(stats,node('p','Last analysed '+state.lastScan.toLocaleString()+' · visible sheet samples','muted'));
  const partial=state.snapshots.filter(s=>s.partial).map(s=>s.name);
  if (partial.length || state.omitted) main.append(node('p',
    (partial.length?'Sampled ranges on '+partial.join(', ')+'. ':'')+
    (state.omitted?state.omitted+' visible sheet(s) beyond the scan limit.':''),'muted'));
  const structure=node('details');
  structure.append(node('summary','Workbook structure'));
  for (const sheet of state.sheets.slice(0,30)) {
    const snapshot=state.snapshots.find(s=>s.id===sheet.id);
    if (!snapshot) {structure.append(node('p',sheet.name+' · '+sheet.visibility,'muted'));continue;}
    const info=profile(snapshot);
    structure.append(node('p',sheet.name+(info.purpose?' · '+info.purpose:''),'meta'),
      node('p',snapshot.tables.length+' table(s) · '+snapshot.names.length+' sheet-scoped name(s) · '+snapshot.chartCount+' chart(s)','muted'),
      node('p',info.columns.slice(0,8).map(c=>c.label+' ('+c.type+')').join(', '),'muted'));
  }
  main.append(structure,node('p',state.monitor?'Monitoring enabled for the first 12 visible sheets when enabled.':'Monitoring paused.','muted'));
  main.append(button('Review recommendations',()=>setTab('Recommendations')));
}

const labels={chart:'Create chart',dashboard:'Create dashboard',style:'Style headings',number:'Change number display',
  formula:'Repair formula',trim:'Trim text whitespace',negative:'Highlight negatives',heatmap:'Apply a colour scale',dataBar:'Add data bars'};
const numberFormats={percent:'0.0%',decimal:'#,##0.00',date:'yyyy-mm-dd',integer:'0'};
function manualPanel(visual) {
  const card=block(node('h3','From the current selection'));
  const options=Object.entries(labels).filter(([kind])=>!['formula','trim'].includes(kind) && (!visual || ['chart','dashboard','heatmap','dataBar'].includes(kind)));
  const kind=select(options,options.some(([v])=>v===state.manualKind)?state.manualKind:options[0][0]);
  kind.addEventListener('change',()=>{state.manualKind=kind.value;render();});
  const option=kind.value==='number'?select(Object.keys(numberFormats).map(k=>[k,k]),'percent'):
    kind.value==='style'?select(Object.keys(palettes).map(k=>[k,k]),'Professional'):
    ['chart','dashboard'].includes(kind.value)?select(Object.keys(chartTypes).map(k=>[k,k]),'column'):null;
  card.append(field('Operation',kind));
  if (option) card.append(field('Option',option));
  card.append(button('Prepare proposal',()=>run(async()=>{
    const created=await makeManual(await selection(),kind.value,option?.value);
    state.tab=visual?'Visualisations':'Recommendations';
    return 'Proposal ready for review on '+findingAddress(created)+'.';
  },'Reading selection…')));
  return card;
}
async function makeManual(s,kind,option) {
  const action={kind,sheetId:s.id,address:s.address};
  if (kind==='number') action.value=numberFormats[option] || numberFormats.percent;
  if (kind==='style') action.value=Object.hasOwn(palettes,option||'')?option:'Professional';
  if (['chart','dashboard'].includes(kind)) {
    action.type=Object.hasOwn(chartTypes,option||'')?option:'column';
    action.title=String(s.values[0]?.[1] || 'Measure').trim().slice(0,90);
    if (kind==='dashboard') action.title+=' overview';
    action.legend=false;
  }
  validate(action);
  if (!available(action)) throw Error('This operation is unavailable in your Excel host.');
  const f={id:JSON.stringify(['manual',s.id,s.address,kind,action.value??action.type??null]), title:labels[kind],
    detail:'Review the proposed change for the selected range.', sheet:s.name, address:s.address, action, manual:true,
    evidence:'Selected range', category:['chart','heatmap','dataBar'].includes(kind)?'Visualisations':kind==='dashboard'?'Analysis':'Formatting'};
  state.dismissed.delete(f.id);
  state.previews.delete(f.id);state.edits.delete(f.id);
  state.findings=[f,...state.findings.filter(x=>x.id!==f.id)];
  return f;
}
function recommendations(visual) {
  main.append(manualPanel(visual));
  if (!visual) {
    const filter=select(['All',...categories].map(c=>[c,c]),state.filter);
    filter.addEventListener('change',()=>{state.filter=filter.value;state.batchReview=false;render();});
    main.append(field('Category',filter));
  }
  const search=node('input');search.type='search';search.value=state.query;search.placeholder='Title, worksheet or range';
  search.addEventListener('change',()=>{state.query=search.value.trim();state.batchReview=false;render();});
  const order=select(['Severity','Confidence','Worksheet','Recommendation type','Expected impact'].map(v=>[v,v]),state.sort);
  order.addEventListener('change',()=>{state.sort=order.value;state.batchReview=false;render();});
  main.append(field('Find recommendations',search),field('Sort by',order));
  const priority=f=>f.priority??0;
  const by={
    Severity:(a,b)=>priority(b)-priority(a),
    Confidence:(a,b)=>Number(b.evidence!=='Business definition required')-Number(a.evidence!=='Business definition required')||priority(b)-priority(a),
    Worksheet:(a,b)=>a.sheet.localeCompare(b.sheet),
    'Recommendation type':(a,b)=>a.category.localeCompare(b.category),
    'Expected impact':(a,b)=>impact(b)-impact(a),
  };
  const list=state.findings.filter(f=>!state.dismissed.has(f.id) && (visual?
    ['Visualisations','Analysis'].includes(f.category):state.filter==='All'||f.category===state.filter))
    .filter(f=>[f.title,f.sheet,f.address,f.detail].join(' ').toLowerCase().includes(state.query.toLowerCase())).sort(by[state.sort]);
  const candidates=list.slice(0,32).filter(f=>state.checked.has(f.id)&&state.previews.has(f.id));
  if (candidates.length>1) {
    const plans=candidates.map(f=>state.previews.get(f.id));
    const collides=plans.some((p,i)=>plans.slice(0,i).some(q=>conflict(p.action,q.action) ||
      p.action.kind==='chart' || q.action.kind==='chart'));
    if (!state.batchReview) main.append(button('Review '+candidates.length+' selected changes',()=>{state.batchReview=true;render();}));
    else main.append(block(node('h3','Batch review'),
      node('p',collides?'This batch contains overlapping ranges or a chart. Apply those changes individually.':
        'Changes run in this order. If one fails, earlier verified changes remain.'),
      ...plans.map((plan,i)=>node('p',candidates[i].title+' · '+plan.snapshot.name+'!'+plan.action.address+' · '+historySummary(plan)))));
    const bulk=button('Approve '+candidates.length+' selected changes',()=>run(async()=>{
    if (collides) throw Error('Resolve the batch conflict before applying.');
    let completed=0;
    try {for (const f of candidates) {await execute(f);completed++;}}
    catch (error) {throw Error(completed+' verified before the batch stopped. '+error.message);}
    return completed+' changes verified.';
  },'Checking and applying selected changes…'),{primary:true,enabled:!collides});
    bulk.dataset.bulk='true';if (state.batchReview) main.append(bulk);
  }
  if (!list.length) main.append(node('p',state.scanned?'No recommendations in this view.':'Run an audit or prepare a proposal from the selection.','muted'));
  list.slice(0,32).forEach(f=>main.append(card(f)));
  if (list.length>32) main.append(node('p','Showing 32 findings. Use the category filter or dismiss reviewed findings to see more.','muted'));
}
async function execute(f) {
  const p=state.previews.get(f.id);
  if (!p || !available(p.action)) throw Error('Preview the operation before approving.');
  try {
    const result=await apply(p);
    state.history.unshift({title:f.title,result:'Verified · '+result.description,recovery:result.recovery,time:new Date(),
      target:p.snapshot.name+'!'+p.action.address,kind:p.action.kind,summary:historySummary(p)});
    // Keep rollback state in memory only and limit the number of retained ranges.
    state.history.slice(25).forEach(item=>delete item.recovery);
    state.dismissed.add(f.id);state.previews.delete(f.id);state.edits.delete(f.id);state.checked.delete(f.id);
  } catch (error) {
    state.history.unshift({title:f.title,result:'Unverified · '+error.message,time:new Date(),target:p.snapshot.name+'!'+p.action.address,kind:p.action.kind});
    state.previews.delete(f.id);state.checked.delete(f.id);
    throw error;
  }
}
function impact(f) {
  const r=rect(f.address);
  return r.rows*r.columns;
}
function historySummary(p) {
  const a=p.action;
  if (a.kind==='formula') {
    const f=formulaPreview(p);
    return String(f.before)+' → '+f.after;
  }
  if (a.kind==='number') return [...new Set(p.snapshot.formats.flat())].slice(0,4).join(', ')+' → '+a.value;
  if (a.kind==='trim') {
    const row=previewRows(p)[0];
    return JSON.stringify(row.before)+' → '+JSON.stringify(row.after);
  }
  if (a.kind==='style') return a.value+' heading style';
  return a.title || labels[a.kind];
}
async function inspectFinding(f,action,withPreview) {
  state.selected=f.id;
  state.previews.delete(f.id);state.checked.delete(f.id);state.batchReview=false;
  const sheetId=action?.sheetId ?? f.sheetId;
  const target=action?.address ?? f.focusAddress ?? f.address;
  const result=await inspectRange(sheetId,target);
  state.inspected.set(f.id,result.snapshot);
  if (withPreview) state.previews.set(f.id,await preview(action));
  return (result.selected?'Selected ':'Inspecting in the pane · ')+result.snapshot.name+'!'+target;
}
function card(f) {
  const chosen=state.selected===f.id;
  const action=state.edits.get(f.id)?.action || state.clarified.get(f.id) || f.action;
  const executable=!!action && available(action);
  const clarification=f.clarification && !action;
  const p=state.previews.get(f.id), inspected=state.inspected.get(f.id);
  const status=executable?(p?'Ready to apply':'Proposal available'):clarification?'Needs review':'Detected';
  const open=()=>run(()=>inspectFinding(f,action,executable),'Reading workbook context…');
  const toggle=button(f.title,()=>{
    if (chosen) {state.selected=null;render();}
    else open();
  });
  toggle.classList.add('finding-title');toggle.setAttribute('aria-expanded',String(chosen));
  const severity=(f.priority??0)>=85?'High priority':(f.priority??0)>=60?'Review':'';
  const card=block(node('div',[f.category,severity,status].filter(Boolean).join(' · '),'meta'),node('h3'),node('div',findingAddress(f),'address'));
  card.querySelector('h3').append(toggle);
  if (chosen) {card.classList.add('selected');card.append(node('p',f.detail),node('p','Evidence · '+f.evidence,'muted'));}
  const actions=node('div',undefined,'actions');
  if (executable) {
    actions.append(button(p?'Refresh preview':'Preview',open));
    if (['chart','dashboard','style'].includes(action.kind))
      actions.append(button(state.edits.has(f.id)?'Close customisation':'Customise',()=>{
        if (state.edits.has(f.id)) {
          state.edits.delete(f.id);state.previews.delete(f.id);state.checked.delete(f.id);state.batchReview=false;render();
        } else run(async()=>{
          const sheets=action.kind==='chart'?await Excel.run(async ctx=>{
            const collection=ctx.workbook.worksheets;
            collection.load('items/id,items/name,items/visibility');await ctx.sync();
            return collection.items.filter(s=>s.visibility==='Visible').map(s=>[s.id,s.name]);
          }):[];
          state.selected=f.id;state.edits.set(f.id,{action:{...action},sheets});
          return '';
        },'Loading options…');
      }));
    if (p && chosen) actions.append(button('Apply changes',()=>run(async()=>{await execute(f);return 'Change verified in Excel.';},'Applying approved change…'),{primary:true}));
  } else {
    actions.append(button('Inspect',open));
    actions.append(button(clarification?'Provide information':'Explain',()=>{
      state.selected=f.id;state.explanation.add(f.id);render();
    }));
    if (action) card.append(node('p','This operation requires a newer Excel API.','muted'));
  }
  actions.append(button(executable?'Decline':'Dismiss',()=>{
    state.dismissed.add(f.id);state.previews.delete(f.id);state.edits.delete(f.id);state.checked.delete(f.id);
    state.inspected.delete(f.id);state.batchReview=false;
    state.history.unshift({title:f.title,result:(executable?'Declined':'Dismissed')+' · '+findingAddress(f),time:new Date()});render();
  }));
  if (chosen && clarification && state.explanation.has(f.id)) {
    const choice=select([['','Choose an answer'],['sum','Yes, add these rows'],['none','No, a total would be misleading']],'');
    const generate=button('Generate proposal',()=>run(async()=>{
      const {metric,type,title}=f.clarification;
      const proposed={kind:'dashboard',sheetId:f.sheetId,address:f.address,type,title,legend:false,confirmedMetric:metric};
      validate(proposed);
      if (!available(proposed)) throw Error('This Excel host does not support dashboard creation.');
      const p=await preview(proposed);
      state.clarified.set(f.id,proposed);state.previews.set(f.id,p);
      return 'Review the proposed dashboard before applying it.';
    },'Preparing dashboard…'),{enabled:false});
    choice.addEventListener('change',()=>{
      generate.disabled=choice.value!=='sum';
      if (choice.value==='none') notice.textContent='No total will be proposed for this metric.';
    });
    card.append(field('Can '+f.clarification.metric+' be added across these rows?',choice),generate);
  } else if (chosen && !executable && state.explanation.has(f.id)) {
    card.append(node('p','Review this evidence against the workbook’s intended calculations or data entry rules. This finding does not contain a proposed workbook edit.','muted'));
  }
  card.append(actions);
  if (chosen && state.edits.has(f.id)) card.append(customisation(f));
  if (chosen && p) {
    card.append(previewPanel(p,f.id));
    const checkbox=node('input');checkbox.type='checkbox';checkbox.checked=state.checked.has(f.id);
    checkbox.disabled=state.busy;
    checkbox.addEventListener('change',()=>{
      if (checkbox.checked) state.checked.add(f.id);else state.checked.delete(f.id);
      state.batchReview=false;render();
    });
    const wrap=node('label',undefined,'check');wrap.append(checkbox,node('span','Select for bulk approval'));card.append(wrap);
  } else if (chosen && inspected) {
    const context=block(node('h3','Workbook context'),node('p',
      inspected.name+'!'+a1(inspected.row,inspected.column,inspected.values.length,inspected.values[0].length),'address'));
    inspected.display.slice(0,4).forEach((row,r)=>context.append(node('p',
      a1(inspected.row+r,inspected.column)+' · '+row.slice(0,5).join('  |  '))));
    card.append(context);
  }
  return card;
}
function customisation(f) {
  const details=node('div',undefined,'details');
  const entry=state.edits.get(f.id), a=entry.action;
  function control(key,label,input) {
    input.disabled=state.busy;
    if (input.type==='checkbox') input.checked=!!a[key];
    else if (input.tagName==='INPUT') input.value=a[key]??'';
    input.addEventListener('change',()=>{
      const value=input.type==='checkbox'?input.checked:input.value;
      if (value===''&&['destination','position','color'].includes(key)) delete a[key];else a[key]=value;
      const showed=state.previews.delete(f.id);state.checked.delete(f.id);
      if (showed) {
        const bulk=main.querySelector('[data-bulk]');if (bulk) bulk.disabled=true;
        for (const control of details.closest('article').querySelectorAll('button,input[type=checkbox]')) {
          if (control.textContent==='Apply changes' || control.type==='checkbox'&&!details.contains(control)) control.disabled=true;
        }
        notice.textContent='Configuration changed. Refresh preview before approving.';
      }
      if (showed || key==='type') render();
    });
    details.append(field(label,input));
  }
  if (a.kind==='style') {
    control('value','Style',select(Object.keys(palettes).map(k=>[k,k]),a.value));
    control('color','Custom fill (hex)',node('input'));return details;
  }
  control('type','Chart type',select(Object.keys(chartTypes).map(k=>[k,k]),a.type));
  control('title','Chart title',node('input'));
  control('address','Data range',node('input'));
  if (a.kind==='chart') {
    control('xLabel','Category axis title',node('input'));control('yLabel','Value axis title',node('input'));
    control('destination','Destination worksheet',select([['','Source worksheet'],...entry.sheets],a.destination||''));
    control('position','Position rectangle (optional)',node('input'));
    if (!['pie','doughnut'].includes(a.type)) control('color','Series colour (optional hex)',node('input'));
  }
  const legend=node('input');legend.type='checkbox';control('legend','Show legend',legend);
  return details;
}
function previewPanel(p,id) {
  const section=node('section',undefined,'details'), {action:a,snapshot:s,data}=p;
  section.append(node('h3','Change preview'),node('p',s.name+'!'+a.address,'address'));
  if (data) {
    section.append(node('p',a.type+' · '+a.title),node('p',(a.xLabel||data.x)+' / '+(a.yLabel||data.y),'muted'));
    if (a.kind==='dashboard') {
      const kpis=node('div',undefined,'metrics');
      for (const [title,value] of [['Total',data.points.reduce((n,[,v])=>n+v,0).toLocaleString()],['Observations',data.points.length]]) {
        const kpi=node('div',undefined,'metric');kpi.append(node('strong',value),node('span',title));kpis.append(kpi);
      }
      section.append(node('p','New worksheet · '+a.title,'meta'),kpis);
    } else section.append(node('p',p.destination+'!'+p.position,'muted'));
    const canvas=node('canvas');
    canvas.setAttribute('role','img');
    canvas.setAttribute('aria-label',a.type+' chart of '+data.y+' by '+data.x);
    section.append(canvas);plot(canvas,data,a);
    const table=node('details');
    table.append(node('summary',a.kind==='dashboard'?'Linked source table':'Chart data'),node('pre',data.points.map(([x,y])=>x+'  '+y).join('\n')));
    section.append(table,node('p','Preview uses source values. Excel controls the final layout and chart rendering.','muted'));
    return section;
  }
  const mode=state.previewMode.get(id)||'Compare';
  const controls=node('div',undefined,'actions');
  for (const option of ['Before','After','Compare']) {
    const toggle=button(option,()=>{state.previewMode.set(id,option);render();});
    toggle.setAttribute('aria-pressed',String(mode===option));controls.append(toggle);
  }
  section.append(controls);
  const comparison=node('div',undefined,'comparison');
  const panels=mode==='Compare'?['Before','After']:[mode];
  let formula;
  if (a.kind==='formula') formula=formulaPreview(p);
  const rows=previewRows(p);
  const values=s.values.flat().filter(v=>typeof v==='number'&&Number.isFinite(v));
  const min=values.reduce((n,v)=>Math.min(n,v),Infinity),max=values.reduce((n,v)=>Math.max(n,v),-Infinity);
  for (const stage of panels) {
    const after=stage==='After';
    const column=node('div',undefined,'preview-column');column.append(node('h4',stage));
    if (formula) {
      column.append(node('code',after?formula.after:formula.before||'Empty'),
        node('p',after?(formula.estimate===null?'Result will be calculated by Excel':'Estimated result · '+formula.estimate):'Current result · '+(formula.current||'Empty'),'muted'));
    } else for (const item of rows) {
      const cell=node('div',undefined,'preview-cell');
      const value=node('span',(after?item.after:item.before)??item.before??'');
      if (a.kind==='trim') value.classList.add('literal-sample');
      cell.append(node('small',item.address),value);
      if (a.kind==='style') {
        const [bold,font,fill]=(after?headingStyle(a):p.currentStyle)??[];
        if (bold) value.style.fontWeight='700';
        if (font) value.style.color=font;
        if (fill) value.style.backgroundColor=fill;
      }
      if (a.kind==='number') cell.append(node('small',after?a.value:item.format));
      const fraction=max===min?0.5:(item.value-min)/(max-min);
      if (after && a.kind==='negative' && item.value<0) value.classList.add('negative-sample');
      if (after && a.kind==='heatmap') value.style.backgroundColor='hsl('+ (208-53*fraction)+' 44% '+(95-36*fraction)+'%)';
      if (after && a.kind==='dataBar') value.style.background='linear-gradient(90deg,#bde1d1 '+(100*fraction)+'%,transparent 0)';
      column.append(cell);
    }
    comparison.append(column);
  }
  section.append(comparison);
  if (formula) {
    const neighbours=node('details');
    neighbours.append(node('summary','Neighbouring formulas'),node('pre',formula.neighbours.map(([cell,value])=>cell+'  '+value).join('\n')));
    section.append(neighbours);
  } else {
    const note=a.kind==='trim'?'Leading and trailing spaces are removed. The original text is preserved in session undo.':a.kind==='number' && a.value==='yyyy-mm-dd'?'Excel will render the date format. Current values are shown in both panels.':
      'Sample of the range. Formatting changes preserve underlying values and formulas.';
    section.append(node('p',note,'muted'));
  }
  return section;
}

function chatView() {
  if (!state.messages.length) main.append(block(node('h3','Ask about your workbook'),
    node('p','Select the relevant cells, then ask for an explanation or propose a change.'),
    button('Run workbook audit',()=>{state.tab='Audit';run(doScan,'Scanning workbook…');},{primary:true})));
  for (const m of state.messages.slice(-18)) {
    const item=block(node('div',m.role==='user'?'You':'SheetGuard AI','meta'),node('p',m.text,'chat'));
    if (m.proposal && !state.dismissed.has(m.proposal.id)) item.append(button('Review proposal',()=>{
      state.filter='All';setTab('Recommendations');}));
    if (m.role==='user') item.classList.add('mine');main.append(item);
  }
  const form=node('form');
  const input=node('textarea');input.maxLength=2000;input.placeholder='What formulas look inconsistent?';input.value=state.draft;
  input.disabled=state.busy;input.addEventListener('input',()=>{state.draft=input.value;});
  form.append(field('Workbook question',input));
  const send=button('Send',()=>{},{primary:true,enabled:state.consent});send.type='submit';
  form.append(send);
  if (!state.consent) form.append(node('p','Enable AI data sharing in Settings to send questions. Local audits remain available.','muted'),
    button('Open Settings',()=>setTab('Settings')));
  form.addEventListener('submit',event=>{
    event.preventDefault();
    const q=state.draft.trim();
    if (q&&state.consent) run(()=>ask(q),'Working on your question…');
  });
  main.append(block(form));
}
async function ask(question) {
  const conversation=state.messages.slice(-6).map(m=>({role:m.role,text:m.text.slice(0,950)}));
  state.messages.push({role:'user',text:question});state.draft='';render();
  try {
    let s;
    try {s=await selection();} catch { /* A question can be answered without a contiguous selection. */ }
    const findings=state.findings.filter(f=>!state.dismissed.has(f.id)).slice(0,12);
    const shorten=x=>typeof x==='string'?x.slice(0,120):x;
    const context=s?{sheet:s.name,range:s.address,totalRows:s.values.length,totalColumns:s.values[0].length,
      values:s.values.slice(0,12).map(r=>r.slice(0,8).map(shorten)),
      formulas:s.formulas.slice(0,12).map(r=>r.slice(0,8).map(x=>typeof x==='string'&&x.startsWith('=')?shorten(x):null))}:undefined;
    const workbook=state.snapshots.slice(0,4).map(sheet=>{
      const summary=profile(sheet);
      return {sheet:shorten(sheet.name),sampled:sheet.partial, purpose:summary.purpose,
        columns:summary.columns.slice(0,6).map(({label,type,formulas,errors})=>({label:shorten(label),type,formulas,errors}))};
    });
    const controller=new AbortController(), timeout=setTimeout(()=>controller.abort(),28000);
    let result;
    try {
      const headers={'Content-Type':'application/json'};
      if (state.accessPassword) {
        const bytes=new TextEncoder().encode('sheetguard:'+state.accessPassword);
        headers.Authorization='Basic '+btoa(String.fromCharCode(...bytes));
      }
      const response=await fetch('/api/agent',{method:'POST',headers,
        body:JSON.stringify({question,selection:context,workbook,findings:findings.map(f=>({title:f.title,detail:f.detail.slice(0,240),
          target:findingAddress(f),actionable:!!f.action&&available(f.action)})),conversation}),signal:controller.signal});
      const data=await response.json();
      if (!response.ok) throw Error(data.error || 'AI request failed.');
      result=data;
    } finally {clearTimeout(timeout);}
    const message={role:'assistant',text:result.reply};
    state.messages.push(message);
    try {
      if (result.intent==='audit') {
        state.tab='Audit';
        return await doScan();
      }
      if (result.finding>=0 && findings[result.finding]?.action) message.proposal=findings[result.finding];
      else if (result.intent!=='none') {
        if (!s) throw Error('Select the affected range to prepare the proposal.');
        message.proposal=await makeManual(s,result.intent,result.option);
      }
    } catch (error) {message.text+='\n\n'+error.message;}
    return '';
  } catch (error) {
    state.messages.push({role:'assistant',text:error.name==='AbortError'?'The AI request timed out. Try again.':error.message});
    return 'AI request ended without a workbook change.';
  } finally {
    if (state.messages.length>60) state.messages.splice(0,state.messages.length-60);
  }
}
function historyView() {
  if (!state.history.length) {main.append(node('p','Approved changes, declined suggestions and verification results will appear here.','muted'));return;}
  state.history.slice(0,50).forEach(item=>{
    const card=block(node('h3',item.title),node('p',item.result),node('p',item.time.toLocaleString(),'muted'));
    if (item.target) card.append(node('p',item.target+' · '+labels[item.kind],'address'));
    if (item.summary) card.append(node('p',item.summary,'muted'));
    if (item.recovery) card.append(button('Undo verified change',()=>{
      if (!window.confirm('Undo this change in Excel?')) return;
      run(async()=>{
        await rollback(item.recovery);delete item.recovery;item.result='Undo verified';
        return 'Undo verified in Excel.';
      },'Checking and undoing change…');
    }));
    main.append(card);
  });
}
function settingsView() {
  const opt=node('input');opt.type='checkbox';opt.checked=state.consent;opt.disabled=state.busy;
  opt.addEventListener('change',()=>{state.consent=opt.checked;render();});
  const label=node('label',undefined,'check');label.append(opt,node('span','Allow AI data sharing for this session'));
  main.append(block(node('h3','AI and workbook data'),label,node('p',
    'Sending a question shares your prompt, up to 12 rows by 8 columns from the selection, up to 12 audit summaries, sampled sheet and column profiles, and recent chat via your SheetGuard host and OpenAI. Workbook scans and writes run inside Excel. Session settings and undo data clear when the pane closes.')));
  const access=node('input');access.type='password';access.maxLength=256;access.value=state.accessPassword;
  access.disabled=state.busy;access.autocomplete='off';
  access.addEventListener('input',()=>{state.accessPassword=access.value;});
  main.append(block(field('Deployment access password',access),
    node('p','For hosted AI, enter the password from your administrator. It clears when you close the pane.','muted')));
  main.append(block(node('h3','Excel capabilities'),...[
    ['Audit', '1.4'], ['Conditional formatting','1.6'],['Charts','1.10'],['Dashboards','1.10'],['Formula and formatting edits','1.13'],
  ].map(([label,version])=>node('p',label+' · '+(supported(version)?'API available':'API unavailable'),'muted'))));
}
render();
if (typeof Office !== 'undefined') {
  Office.onReady(info=>{
    state.ready=info.host===Office.HostType.Excel;
    connection.textContent=state.ready?'Connected to Excel':'Excel host required';
    render();
    if (state.ready) run(doScan,'Scanning workbook…');
  }).catch(()=>{connection.textContent='Excel connection failed';});
} else connection.textContent='Office.js unavailable';
