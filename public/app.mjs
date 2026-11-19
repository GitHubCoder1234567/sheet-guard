import {formula, formulaReferences, datasetColumns, summariseBy, compareDatasets, a1, rect, previewRows, formulaPreview, headingStyle, profile, audit, categories, chartTypes, palettes, conflict, validate} from './engine.mjs';
import {visibleWorksheetByName, selectEvidenceRange, activeWorksheet, readDataset, readContext, selectionAnchor, scan, watchSheets, selection, inspectRange, preview, apply, rollback, supported} from './excel.mjs';
import {plot} from './plot.mjs';
import {aggregationEvidence, sameEvidence, evidenceIntent, evidenceGroup, foldLabel} from './evidence.mjs';
import {calculationGoal, calculatesGoal, datasetAnswer, executeTool, explicitAudit, localQuestion, rankedAnalysis} from './agent.mjs';

const nav = document.querySelector('#nav'), main = document.querySelector('#main');
const notice = document.querySelector('#notice'), connection = document.querySelector('#connection');
const tabs = ['Audit','Recommendations','Visualisations','History','Settings'];
const state = {
  sessionId:globalThis.crypto?.randomUUID?.()??'forma-'+Date.now(), tab:'Chat', workspaceMode:null, menuOpen:false, ready:false, busy:false, findings:[], snapshots:[], dismissed:new Set(),
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
function setTab(tab) {state.tab=tab;state.menuOpen=false;render();}
function render() {
  const menu=node('details',undefined,'workspace-menu');
  menu.open=state.menuOpen;
  const summary=node('summary','Menu');
  summary.setAttribute('aria-label','Workspace menu');
  menu.append(summary);
  const links=node('div',undefined,'menu-items');
  for (const tab of tabs) {
    const b=button(tab,()=>setTab(tab),{during:true});
    if (tab===state.tab) b.setAttribute('aria-current','page');
    links.append(b);
  }
  menu.append(links);
  menu.addEventListener('toggle',()=>{if (menu.isConnected) state.menuOpen=menu.open;});
  nav.replaceChildren(menu);
  main.replaceChildren();
  if (!state.ready) {
    main.append(block(node('p','Open Forma inside Microsoft Excel to work with a workbook.')));
    return;
  }
  if (state.tab!=='Chat') main.append(
    button('Back to workspace',()=>setTab('Chat'),{during:true}),node('h2',state.tab));
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
  if (a.kind==='pivot') return supported('1.8');
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
  formula:'Repair formula',normalizeLabel:'Standardise category label',pivot:'Create PivotTable',nonnegative:'Prevent negative values',trim:'Trim text whitespace',negative:'Highlight negatives',heatmap:'Apply a colour scale',dataBar:'Add data bars'};
const numberFormats={percent:'0.0%',decimal:'#,##0.00',date:'yyyy-mm-dd',integer:'0'};
function manualPanel(visual) {
  const card=block(node('h3','From the current selection'));
  const options=Object.entries(labels).filter(([kind])=>!['formula','trim','pivot','nonnegative','normalizeLabel'].includes(kind) && (!visual || ['chart','dashboard','heatmap','dataBar'].includes(kind)));
  const kind=select(options,options.some(([v])=>v===state.manualKind)?state.manualKind:options[0][0]);
  kind.addEventListener('change',()=>{state.manualKind=kind.value;state.toolsOpen=kind.closest('details')?.open;render();});
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
function recommendations(visual,inline=false) {
  if (!inline) {
    const tools=node('details',undefined,'tools');
    tools.open=!!state.toolsOpen;
    tools.addEventListener('toggle',()=>{if(tools.isConnected) state.toolsOpen=tools.open;});
    tools.append(node('summary','Prepare a change from the selection'),manualPanel(visual));main.append(tools);
  }
  if (!visual && !inline) {
    const filter=select(['All',...categories].map(c=>[c,c]),state.filter);
    filter.addEventListener('change',()=>{state.filter=filter.value;state.batchReview=false;render();});
    main.append(field('Category',filter));
  }
  const search=node('input');search.type='search';search.value=state.query;search.placeholder='Title, worksheet or range';
  search.addEventListener('change',()=>{state.query=search.value.trim();state.batchReview=false;render();});
  const order=select(['Severity','Confidence','Worksheet','Recommendation type','Expected impact'].map(v=>[v,v]),state.sort);
  order.addEventListener('change',()=>{state.sort=order.value;state.batchReview=false;render();});
  if (!inline) main.append(field('Find recommendations',search),field('Sort by',order));
  const priority=f=>f.priority??0;
  const by={
    Severity:(a,b)=>priority(b)-priority(a),
    Confidence:(a,b)=>Number(b.evidence!=='Business definition required')-Number(a.evidence!=='Business definition required')||priority(b)-priority(a),
    Worksheet:(a,b)=>a.sheet.localeCompare(b.sheet),
    'Recommendation type':(a,b)=>a.category.localeCompare(b.category),
    'Expected impact':(a,b)=>impact(b)-impact(a),
  };
  const list=state.findings.filter(f=>!state.dismissed.has(f.id) && (visual?
    ['Visualisations','Analysis'].includes(f.category):inline||state.filter==='All'||f.category===state.filter))
    .filter(f=>!inline || state.workspaceMode!=='clean' || f.category==='Data Quality')
    .filter(f=>inline || [f.title,f.sheet,f.address,f.detail].join(' ').toLowerCase().includes(state.query.toLowerCase())).sort(by[state.sort]);
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
  if (a.kind==='pivot') return 'Sum of '+a.measure+' by '+a.dimension+' · new worksheet';
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
  if (chosen) {
    card.classList.add('selected');
    const simple=f.action?.kind==='formula'?'Nearby rows use a different calculation. Compare this cell with its neighbours before applying a correction.':f.plain||f.detail;
    card.append(node('p',simple));
    if (f.action?.kind==='formula') card.append(node('p','This is a possible inconsistency. The workbook may intentionally use a different rule here.','muted'));
    const further=node('details',undefined,'explanation');
    const technical=node('details');
    technical.append(node('summary','Technical evidence'),node('p',f.detail));
    further.append(node('summary','Explain further'),node('p','Affected range · '+findingAddress(f)),
      node('p','Evidence · '+f.evidence),technical);
    card.append(further);
    if (f.matchRows) {
      const table=node('details');
      table.append(node('summary','Compare matching rows'));
      for(const row of f.matchRows) table.append(node('p',row.address+' · '+row.values.join('  |  ')));
      card.append(table);
    }
  }
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
  if (a.kind==='normalizeLabel') {
    section.append(node('p','One literal category cell will change. Sales values, formulas and formatting are preserved.'),
      node('p','Before · '+JSON.stringify(a.before)),node('p','After · '+JSON.stringify(a.value)),
      node('p','Changing a label can regroup this record in reports. You can undo the verified change from History.','muted'));
    return section;
  }
  if (a.kind==='nonnegative') {
    section.append(node('p','Keep current values. Future entries must be zero or greater. Blank cells remain allowed.'),
      node('p','Excel will display a Stop alert when a typed entry fails this rule. No existing rule will be replaced.','muted'));
    return section;
  }
  if (p.report) {
    const report=p.report;
    section.append(node('p','Native PivotTable · new worksheet'),node('p','Sum of '+report.measure+' by '+report.dimension),
      node('p',report.count+' source records · '+report.groups.length+' groups','muted'));
    for (const group of report.groups.slice(0,12)) section.append(node('p',group.label+' · '+group.sum.toLocaleString()));
    if (report.groups.length>12) section.append(node('p','Preview shows 12 groups. The report will include them all.','muted'));
    return section;
  }
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

function launchReview(mode) {
  state.workspaceMode=mode;
  run(doScan,'Reviewing workbook…');
}

function chatView() {
  const form=node('form',undefined,'composer');
  const input=node('textarea');input.maxLength=2000;
  input.placeholder='Ask Forma anything about your spreadsheet…';input.value=state.draft;
  input.disabled=state.busy;input.addEventListener('input',()=>{state.draft=input.value;});
  form.append(field('Workbook question',input));
  const send=button('Send',()=>{},{primary:true});send.type='submit';
  form.append(send);
  if (state.chatError && !state.busy) {
    const error=node('p',state.chatError,'chat-error');
    error.setAttribute('role','alert');form.append(error);
  }
  form.addEventListener('submit',event=>{
    event.preventDefault();
    const q=state.draft.trim();
    if (!q) return;
    if (/^(?:review|check|audit|scan)(?: my| this| the)? (?:spreadsheet|workbook|sheet)[.!]?$/i.test(q)) {
      state.draft='';launchReview('review');return;
    }
    if (/^(?:clean|clean and organise)(?: my| this| the)? (?:dataset|data|table|spreadsheet)[.!]?$/i.test(q)) {
      state.draft='';launchReview('clean');return;
    }
    if (/^compare (?:these |two )?(?:datasets|ranges|tables|worksheets)[.!]?$/i.test(q)) {
      state.draft='';state.workspaceMode='compare';state.comparison={};render();return;
    }
    if (/^(?:prevent negative (?:values|numbers|quantities)|require non-negative (?:values|numbers))[.!]?$/i.test(q)) {
      state.draft='';run(()=>prepareValidation(q),'Preparing validation…');return;
    }
    if (/^create (?:a )?pivot\s*table for total .+ by .+[.!]?$/i.test(q)) {
      state.draft='';run(()=>preparePivot(q),'Preparing report…');return;
    }
    if (/^undo (?:it|that|the last change)[.!]?$/i.test(q)) {
      state.draft='';
      state.chatError='';
      setTab('History');
      return;
    }
    const evidenceRequest=evidenceIntent(q);
    const local=localQuestion(q);
    if (evidenceRequest && (!local || (state.evidence && evidenceRequest==='normalise'))) {
      run(()=>submitQuestion(q,()=>answerEvidence(q,evidenceRequest)),'Checking source evidence…');return;
    }
    if (local) {
      run(()=>submitQuestion(q,()=>answerLocal(q,local)),'Reading workbook…');return;
    }
    if (!state.consent) {
      state.chatError='Enable AI data sharing in Settings to send this question. Local reviews and calculations stay in Excel.';
      render();
      return;
    }
    run(()=>submitQuestion(q,conversation=>ask(q,conversation)),'Working on your question…');
  });
  main.append(form);
  if (state.datasetController) main.append(button('Cancel reading',()=>state.datasetController.abort(),{during:true}));
  const suggestions=node('div',undefined,'actions suggestions');
  suggestions.append(button('Review my spreadsheet',()=>launchReview('review')));
  const relevant=state.findings.some(f=>f.category==='Data Quality');
  suggestions.append(button(relevant?'Clean and organise':'Compare datasets',()=>{
    if (relevant) launchReview('clean');
    else {state.workspaceMode='compare';state.comparison={};render();}
  }));
  suggestions.append(button('Explore insights',()=>launchReview('insights')));
  main.append(suggestions);
  for (const m of state.messages.slice(-18)) {
    const item=block(node('div',m.role==='user'?'You':'Forma','meta'),node('p',m.text,'chat'));
    if (m.report) item.append(rankingTable(m));
    if (m.explanation) item.append(node('p',m.explanation,'muted'));
    if (m.rows) item.append(evidenceRows(m));
    if (m.proposal && !state.dismissed.has(m.proposal.id)) item.append(button('Review proposal',()=>{
      state.filter='All';setTab('Recommendations');}));
    if (m.role==='user') item.classList.add('mine');main.append(item);
  }
  if (state.workspaceMode==='compare') main.append(comparisonPanel());
  else if (state.workspaceMode) {
    main.append(node('h2',{review:'Workbook review',clean:'Cleaning opportunities',insights:'Workbook insights'}[state.workspaceMode]));
    if (state.scanController) main.append(button('Cancel scan',()=>state.scanController.abort(),{during:true}));
    if (state.lastScan) {
      const cells=state.snapshots.reduce((total,s)=>total+s.values.length*s.values[0].length,0);
      main.append(node('p',cells.toLocaleString()+' cells reviewed · '+state.snapshots.length+' visible sheets'+
        (state.omitted||state.snapshots.some(s=>s.partial)?' · partial coverage':''),'muted'));
      recommendations(state.workspaceMode==='insights',true);
    }
  }
}
function evidenceRows(message) {
  const rows=message.rows, page=message.rowPage??0, size=10;
  const section=node('section',undefined,'details');
  section.append(node('h4','Source records · '+rows.length));
  for (const row of rows.slice(page*size,(page+1)*size)) {
    const label=message.sheet+'!'+row.groupCell+' · '+row.label+' / '+message.sheet+'!'+row.measureCell+' = '+row.value;
    section.append(node('p',label+(row.identifier===null?'':' · '+row.identifierCell+' '+row.identifier)));
    if (row.displayed!==null && row.displayed!==String(row.value))
      section.append(node('p','Displayed in Excel: '+row.displayed,'muted'));
    section.append(
      button('Select '+row.measureCell,()=>run(async()=>{
        const previous=await selectionAnchor();
        await selectEvidenceRange(message.sheetId,row.measureCell);
        state.evidenceBack=previous;
        return 'Selected '+message.sheet+'!'+row.measureCell+'.';
      },'Navigating to source…')));
  }
  if (rows.length>size) {
    const control=node('div',undefined,'actions');
    for (const [label,dir] of [['Previous records',-1],['Next records',1]])
      control.append(button(label,()=>{message.rowPage=page+dir;render();},
        {enabled:dir<0?page>0:(page+1)*size<rows.length}));
    section.append(control);
  }
  if (state.evidenceBack) section.append(button('Return to previous selection',()=>run(async()=>{
    const previous=state.evidenceBack;
    await selectEvidenceRange(previous.id,previous.address);
    state.evidenceBack=null;
  },'Restoring selection…')));
  return section;
}

async function answerEvidence(question,kind,exactGroup) {
  const prior=state.evidence;
  if (!prior || prior.sessionId!==state.sessionId) throw Error('Calculate the category totals first so Forma can identify their source rows.');
  const {source,fields}=prior;
  const data=await captureDataset({sheetId:source.sheetId,address:source.address},false,true);
  const current=aggregationEvidence(data,fields,{sessionId:state.sessionId,normalised:prior.grouping!=='exact label'});
  if (!sameEvidence(prior,current)) {
    state.evidence=null;
    throw Error('The source cells changed since the previous calculation. Run the ranking again before using its totals or proposing edits.');
  }
  if (kind==='regroup') {
    const report=rankedAnalysis({...fields,sessionId:state.sessionId,normalised:true,operation:'highest'},data);
    state.evidence=report.evidence;state.evidenceGroup=report.groups[0].label;
    state.messages.push({role:'assistant',text:report.text,report});
    return;
  }
  const add=(text,rows,proposal)=>state.messages.push({role:'assistant',text,rows,proposal,sheet:source.sheet,sheetId:source.sheetId});
  if (kind==='records') {
    const group=exactGroup===undefined?evidenceGroup(prior,question,state.evidenceGroup):prior.groups.find(g=>g.label===exactGroup);
    if (!group) throw Error('That category is no longer available in the saved result.');
    state.evidenceGroup=group.label;
    const sum=group.records.length<=12?group.records.map(x=>String(x.value)).join(' + ')+' = '+group.sum:
      group.records.length+' numeric source values sum to '+group.sum;
    add(group.label+' · '+sum+'. These records contribute to the aggregate.',group.records);
    return;
  }
  const focus=typeof state.evidenceGroup==='string'?foldLabel(state.evidenceGroup):null;
  const candidates=[...prior.variants].filter(([,labels])=>labels.size>1);
  const match=candidates.find(([key])=>key===focus)|| (candidates.length===1?candidates[0]:null);
  if (!match) throw Error('One category with a case or whitespace variant could not be identified. Ask for its exact total first.');
  const [,labels]=match, rows=[...labels.values()].flat();
  const folded=aggregationEvidence(data,fields,{sessionId:state.sessionId,normalised:true});
  const combined=folded.groups.find(g=>g.label===match[0]);
  if (kind==='variants') {
    add('Exact grouping keeps '+[...labels.keys()].map(x=>JSON.stringify(x)).join(' and ')+
      ' separate. If you opt into case-insensitive grouping with normalised whitespace, their combined total is '+combined.sum+'.',rows);
    return;
  }
  const [canonical,...changes]=[...labels];
  const variants=changes.flatMap(([,records])=>records);
  if (variants.length!==1) throw Error('This proposal supports one affected cell. More than one variant cell was found; review the source records first.');
  const record=variants[0];
  const action={kind:'normalizeLabel',sheetId:source.sheetId,address:record.groupCell,before:record.label,value:canonical[0]};
  const f={id:JSON.stringify(['normalise',source.sheetId,record.groupCell,record.label,canonical[0]]),
    title:'Standardise category label',detail:'Review the case or whitespace change for this one record.',
    sheet:source.sheet,address:record.groupCell,action,manual:true,category:'Data Quality',evidence:'Exact source records'};
  await preview(action);
  state.findings=[f,...state.findings.filter(x=>x.id!==f.id)];state.dismissed.delete(f.id);
  add('Proposed '+source.sheet+'!'+record.groupCell+' · '+JSON.stringify(record.label)+' → '+JSON.stringify(canonical[0])+
    '. Review and approve the change to apply it.',rows,f);
}

function rankingTable(message) {
  const report=message.report, pageSize=25;
  const pages=Math.ceil(report.groups.length/pageSize), page=message.page??0;
  const table=node('table',undefined,'ranking-table'), caption=node('caption','Ranked totals · '+report.groups.length.toLocaleString()+' groups');
  const header=node('thead'), columns=node('tr');
  for (const label of ['Rank',report.dimension,report.measure,'Rows']) columns.append(node('th',label));
  header.append(columns);table.append(caption,header);
  const body=node('tbody');
  for (const group of report.groups.slice(page*pageSize,(page+1)*pageSize)) {
    const row=node('tr');
    const category=node('td');
    category.append(button(String(group.label),()=>run(async()=>{
      const previous=state.evidence;
      state.evidence=report.evidence;
      try {await answerEvidence('', 'records',group.label);}
      catch(error) {
        if (state.evidence!==null) state.evidence=previous;
        throw error;
      }
    },'Checking source records…')));
    row.append(node('td',group.rank),category,node('td',group.sum.toLocaleString(undefined,{maximumFractionDigits:10})),
      node('td',group.records));
    body.append(row);
  }
  table.append(body);
  const wrap=node('div',undefined,'ranking-wrap');
  wrap.append(table,node('p','Select a category to see its source records.','muted'));
  if (pages>1) {
    const controls=node('div',undefined,'actions');
    const change=delta=>{message.page=page+delta;render();};
    controls.append(button('Previous',()=>change(-1),{enabled:page>0}),
      node('span',(page+1)+' / '+pages,'muted'),button('Next',()=>change(1),{enabled:page+1<pages}));
    wrap.append(controls);
  }
  return wrap;
}
async function captureDataset(target,expand=false,evidence=false) {
  state.datasetController=new AbortController();
  render();
  try {
    return await readDataset(state.datasetController.signal,(read,total)=>{
      notice.textContent='Read '+read.toLocaleString()+' of '+total.toLocaleString()+' rows.';
    },target,expand,evidence);
  } finally {state.datasetController=null;}
}
function comparisonPanel() {
  const view=state.comparison;
  const panel=block(node('h2','Compare datasets'),node('p','In Excel, select the first range including its headers, capture it, then repeat for the second range.'));
  for (const [key,label] of [['first','First'],['second','Second']]) {
    panel.append(button('Capture '+label.toLowerCase()+' selection',()=>run(async()=>{
      const dataset=await captureDataset();
      datasetColumns(dataset.values);
      view[key]=dataset;view.result=null;
      return label+' selection captured.';
    },'Reading selection…')));
    if (view[key]) panel.append(node('p',view[key].sheet+'!'+view[key].address+' · '+(view[key].values.length-1)+' data rows','address'));
  }
  if (!view.first || !view.second) return panel;
  const right=new Set(datasetColumns(view.second.values).map(v=>v.toLowerCase()));
  const options=datasetColumns(view.first.values).filter(v=>right.has(v.toLowerCase()));
  if (!options.length) {panel.append(node('p','The captures need a shared identifier column header.'));return panel;}
  const key=select([['','Choose a matching column'],...options.map(v=>[v,v])],view.key??'');
  key.addEventListener('change',()=>{view.key=key.value;view.result=null;render();});
  panel.append(field('Match records by',key),button('Compare records',()=>run(async()=>{
    view.result=compareDatasets(view.first,view.second,view.key);return 'Comparison ready.';
  },'Comparing records…'),{enabled:!!options.includes(view.key),primary:true}));
  if (!view.result) return panel;
  const r=view.result;
  panel.append(node('p',r.added+' new · '+r.removed+' missing · '+r.changed+' changed · '+r.unchanged+' unchanged'));
  panel.append(node('p',r.ambiguous+' repeated key values excluded · '+r.missing+' rows with unusable keys · '+r.ignoredErrors+' matched records contain errors','muted'));
  if (r.firstOnly.length||r.secondOnly.length) panel.append(node('p','Columns only in first: '+(r.firstOnly.join(', ')||'none')+
    '. Only in second: '+(r.secondOnly.join(', ')||'none')+'.','muted'));
  const detail=node('details');detail.append(node('summary','Sample changes'));
  for (const row of r.samples) {
    detail.append(node('p',row.type+' · '+row.key,'meta'));
    for (const cell of row.changes??[]) detail.append(node('p',cell.name+': '+cell.before+' → '+cell.after+
      (cell.delta===undefined?'':' (difference '+cell.delta.toLocaleString()+
        (cell.percent===undefined?'':', '+new Intl.NumberFormat(undefined,{style:'percent',maximumFractionDigits:2}).format(cell.percent))+')')));
  }
  panel.append(detail);return panel;
}
async function prepareValidation(question) {
  const s=await selection();
  const action={kind:'nonnegative',sheetId:s.id,address:s.address};
  await preview(action);
  const f={id:JSON.stringify(['manual',s.id,s.address,action.kind]),manual:true,category:'Data Quality',
    title:'Prevent negative values',detail:'Require numbers greater than or equal to zero in the selected cells.',
    evidence:'Requested entry rule',sheet:s.name,address:s.address,action};
  state.findings=[f,...state.findings.filter(x=>x.id!==f.id)];state.dismissed.delete(f.id);
  state.messages.push({role:'user',text:question},{role:'assistant',text:'The validation proposal is ready to review.',proposal:f});
}
async function preparePivot(question) {
  const [,measure,dimension]=/^create (?:a )?pivot\s*table for total (.+) by (.+)$/i.exec(question.replace(/[.!]$/,''));
  if (!supported('1.8')) throw Error('This Excel host lacks native PivotTables. Ask “total '+measure+' by '+dimension+'” for an on-screen summary.');
  const source=await selection();
  const action={kind:'pivot',sheetId:source.id,address:source.address,measure,dimension};
  const f={id:JSON.stringify(['manual',source.id,source.address,'pivot',measure,dimension]),manual:true,
    category:'Analysis', title:'PivotTable · '+measure+' by '+dimension,detail:'Summarise the selected data on a new worksheet.',
    sheet:source.name,address:source.address,evidence:'Selected data and requested fields',action};
  await preview(action);
  state.findings=[f,...state.findings.filter(x=>x.id!==f.id)];state.dismissed.delete(f.id);
  state.messages.push({role:'user',text:question},{role:'assistant',text:'The PivotTable proposal is ready to review.',proposal:f});
  return '';
}

async function submitQuestion(question,work) {
  state.chatError='';
  const conversation=state.messages.filter(m=>!m.failed&&!m.report&&!m.rows).slice(-6).map(m=>({role:m.role,text:m.text.slice(0,950)}));
  const previous=state.messages.at(-1);
  const turn=previous?.failed && previous.text===question?previous:{role:'user',text:question};
  if (turn!==previous) state.messages.push(turn);
  try {
    const status=await work(conversation);
    turn.failed=false;
    state.draft='';
    return status;
  } catch(error) {
    turn.failed=true;
    state.draft=question;
    state.chatError=error.name==='AbortError'?'The request was interrupted or timed out. Try again.':error.message;
  } finally {
    if (state.messages.length>60) state.messages.splice(0,state.messages.length-60);
  }
}
async function answerLocal(question,intent) {
  const query=question.replace(/[.!]$/,'');
  const total=/^(?:show me |calculate )?total (.+) by (.+)$/i.exec(query);
  let reply, report;
  if (intent.kind==='sheet') {
    const sheet=await activeWorksheet();
    reply='The current worksheet is “'+sheet.name+'”.';
  } else if (intent.kind==='errors') {
    const sheet=await activeWorksheet();
    const {snapshots}=await scan(undefined,undefined,new Set([sheet.id]),true);
    const snapshot=snapshots[0];
    const errors=[];
    if (snapshot) for (let r=0;r<snapshot.values.length;r++) for (let c=0;c<snapshot.values[r].length;c++)
      if (snapshot.types[r][c]==='Error') errors.push(a1(snapshot.row+r,snapshot.column+c)+' '+snapshot.display[r][c]);
    const scope=snapshot?sheet.name+'!'+a1(snapshot.row,snapshot.column,snapshot.values.length,snapshot.values[0].length):sheet.name;
    reply=(errors.length?'Found '+errors.length+' Excel error values in '+scope+': '+errors.slice(0,12).join(', ')+
      (errors.length>12?' (first 12 shown).':'.'):'No Excel error values found in '+scope+'.')+
      (snapshot?.partial?' This was a bounded sample.':'');
  } else if (['count','columns','rank'].includes(intent.kind)) {
    const sheet=intent.sheetName?await visibleWorksheetByName(intent.sheetName):intent.worksheet?await activeWorksheet():null;
    const data=await captureDataset(sheet?{sheetId:sheet.id,used:true}:undefined,true,intent.kind==='rank');
    if (intent.kind==='rank') {
      report=rankedAnalysis({...intent,sessionId:state.sessionId},data);
      state.evidence=report.evidence;
      state.evidenceGroup=intent.operation==='lowest'?report.groups.at(-1).label:report.groups[0].label;
      reply=report.text;
    } else reply=datasetAnswer(intent,data);
  } else if (total) {
    const data=await captureDataset(undefined,true), result=summariseBy(data,total[1],total[2]);
    const top=result.groups.slice(0,12).map(({label,sum,records})=>
      label+': '+sum.toLocaleString(undefined,{maximumFractionDigits:8})+' ('+records+' rows)');
    reply='Total '+result.measure+' by '+result.dimension+'\n\n'+(top.join('\n')||'No eligible numeric rows.')+
      '\n\nSource: '+data.sheet+'!'+data.address+'. '+result.count+' included, '+result.omitted+' omitted.'+
      (result.groups.length>12?' Showing the 12 largest of '+result.groups.length+' groups.':'');
  } else {
    const s=await selection();
    if (s.values.length!==1 || s.values[0].length!==1) throw Error('Select one formula cell to explain it.');
    if (!formula(s,0,0)) throw Error('Select a cell containing a formula.');
    const expression=s.formulas[0][0];
    const description=/^\s*=\s*XLOOKUP\(/i.test(expression)?
      'XLOOKUP searches for a value in one range and returns the corresponding value from another. A missing match can return #N/A unless the formula supplies a fallback.':
      'Excel calculates this cell using the formula below. The references listed are the cells and ranges Forma can identify.';
    const refs=formulaReferences(expression.slice(0,4096)).slice(0,15).map(x=>(x.sheet?x.sheet+'!':'')+x.first+(x.last?':'+x.last:''));
    reply=s.name+'!'+s.address+'\n'+description+'\n\n'+expression+'\nCurrent result: '+s.display[0][0]+
      (refs.length?'\nReferences: '+refs.join(', '):'');
  }
  const message={role:'assistant',text:reply,report};
  state.messages.push(message);
  if (report && intent.explainVerified && state.consent) {
    notice.textContent='Calculations verified. Preparing an optional explanation…';
    render();
    const {operation,grouping,coverage,counts}=report.evidence;
    try {
      const answer=await requestAgent({question:'Explain the verified grouping and summation method and its limitations.',
        toolResult:{verifiedCalculation:{operation,grouping,coverage,counts,groupCount:report.groups.length}}});
      if (answer.intent!=='none' || answer.finding!==-1 || (answer.tool && answer.tool!=='none') ||
          /[0-9]/.test(answer.reply) || answer.reply.length>600)
        throw Error('Groq returned an explanation that could not be verified.');
      message.explanation='AI explanation · '+answer.reply;
    } catch(error) {
      message.explanation='The verified totals and source cells are above. AI explanation unavailable. '+
        (error.name==='AbortError'?'The explanation timed out.':error.message);
    }
  }
  return '';
}
async function requestAgent(payload) {
  const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),22000);
  const headers={'Content-Type':'application/json'};
  if (state.accessPassword) {
    const bytes=new TextEncoder().encode('sheetguard:'+state.accessPassword);
    headers.Authorization='Basic '+btoa(String.fromCharCode(...bytes));
  }
  try {
    let response;
    try {
      response=await fetch('/api/agent',{method:'POST',headers,body:JSON.stringify(payload),signal:controller.signal});
    } catch (error) {
      if (error.name==='AbortError') throw error;
      throw Error('Forma could not reach its AI service. Check your connection and try again.');
    }
    const data=await response.json().catch(()=>null);
    if (!response.ok) {
      if (data && typeof data.error==='string' && data.error.trim())
        throw Error(data.error+(typeof data.reference==='string' && /^[0-9a-f]{12}$/.test(data.reference)?
          ' Reference '+data.reference+'.':''));
      if (response.status===401) throw Error('Enter the deployment access password in Forma Settings.');
      if (response.status===429) throw Error('The AI service is rate limited. Wait a moment and try again.');
      throw Error('The AI service returned HTTP '+response.status+'. Try again shortly.');
    }
    if (!data || typeof data.reply!=='string' || !data.reply.trim() ||
        !['none','audit','style','number','chart','dashboard','negative','heatmap','dataBar'].includes(data.intent) ||
        typeof data.option!=='string' || !Number.isInteger(data.finding) ||
        (data.tool!==undefined && typeof data.tool!=='string'))
      throw Error('The AI service returned an invalid answer. Try asking a narrower question.');
    if (payload.toolResult && data.tool && data.tool!=='none' &&
        !(payload.toolResult.calculationPending===true && ['AggregateData','AnalyseColumn'].includes(data.tool))) throw Error('The AI requested another tool after receiving the workbook evidence.');
    return data;
  } finally {clearTimeout(timeout);}
}

async function ask(question,conversation) {
    let anchor,s;
    try {
      anchor=await selectionAnchor();
      if (anchor.rows*anchor.columns<=2400) s=await readContext({sheetId:anchor.id,address:anchor.address});
      if (s) s={...s,address:anchor.address};
    } catch { /* Workbook metadata remains usable without a selection. */ }
    const findings=state.findings.filter(f=>!state.dismissed.has(f.id)).slice(0,12);
    const shorten=x=>typeof x==='string'?x.slice(0,120):x;
    const context=s?{sheet:s.name,range:s.address,totalRows:s.values.length,totalColumns:s.values[0].length,
      values:s.values.slice(0,12).map(r=>r.slice(0,8).map(shorten)),
      formulas:s.formulas.slice(0,12).map(r=>r.slice(0,8).map(x=>typeof x==='string'&&x.startsWith('=')?shorten(x):null))}:
      anchor?{sheet:anchor.name,range:anchor.address,totalRows:anchor.rows,totalColumns:anchor.columns}:undefined;
    const workbook=state.snapshots.slice(0,4).map(sheet=>{
      const summary=profile(sheet);
      return {sheet:shorten(sheet.name),sampled:sheet.partial, purpose:summary.purpose,
        columns:summary.columns.slice(0,6).map(({label,type,formulas,errors})=>({label:shorten(label),type,formulas,errors}))};
    });
    const payload={question,selection:context,workbook,findings:findings.map(f=>({title:f.title,detail:f.detail.slice(0,240),
      target:findingAddress(f),actionable:!!f.action&&available(f.action)})),conversation};
    const goal=calculationGoal(question);
    if (goal==='unsupported') throw Error('Forma cannot verify that statistic with its current calculation tools.');
    let result=await requestAgent(payload);
    let routedProposal;
    const needsCalculation=!!goal;
    const requested=new Set();
    for (let step=0;step<2 && result.tool && result.tool!=='none';step++) {
      if (requested.has(result.tool)) throw Error('The assistant repeated its workbook request. Select a specific range and try again.');
      requested.add(result.tool);
      if (needsCalculation && !['InspectWorkbook','InspectWorksheet','InspectRange'].includes(result.tool) &&
          !calculatesGoal(goal,result))
        throw Error('No calculation was performed. Select the relevant data and specify the measure or grouping column.');
      let outcome;
      try {outcome=await executeTool(result,{sheets:state.sheets,snapshots:state.snapshots,
        selection:anchor?{...s,...anchor}:null,readRange:readContext,
        detailed:goal==='ranking'||needsCalculation&&/\b(?:rank|all|every|explain|source)\b/i.test(question),
        rankingOperation:goal==='ranking'&&/\blowest\b/i.test(question)?'lowest':undefined,
        sessionId:state.sessionId,readDataset:(target,evidence)=>{
          if (!target&&!anchor) throw Error('Select the dataset, including its header.');
          return captureDataset(target??{sheetId:anchor.id,address:anchor.address},true,evidence);
        }});
      } catch(error) {
        if (error.name==='AbortError') throw error;
        const operation=result.tool.replace(/([a-z])([A-Z])/g,'$1 $2');
        throw Error(operation+' failed. '+error.message);
      }
      if (outcome.local) {
        if (outcome.report) {
          state.evidence=outcome.report.evidence;
          state.evidenceGroup=outcome.report.groups[0].label;
        }
        state.messages.push({role:'assistant',text:outcome.local,report:outcome.report});
        return '';
      }
      if (outcome.compare) {
        state.workspaceMode='compare';state.comparison={};
        state.messages.push({role:'assistant',text:'Capture the two ranges below and choose their matching column.'});
        return '';
      }
      routedProposal=outcome.proposal;
      if (routedProposal) {
        if (result.address || result.sheet)
          throw Error('Select the intended range directly in Excel before requesting a change.');
        if (!s) throw Error('Select a smaller editable range before requesting a proposal.');
      }
      if (outcome.evidence) {
        result=await requestAgent({question,toolResult:{...outcome.evidence,calculationPending:needsCalculation&&step===0}});
        continue;
      }
      break;
    }
    if (needsCalculation)
      throw Error('No calculation was performed. Select the source data and include the exact column headings in your question.');
    if (result.intent==='audit' && !explicitAudit(question))
      throw Error('Forma could not match that request to a workbook action. Describe the analysis or change you want, or name the relevant range.');
    const message={role:'assistant',text:routedProposal?'Review the proposed change before applying it.':result.reply.trim()};
    if (routedProposal || (result.finding<0 && result.intent!=='none' && result.intent!=='audit')) {
      if (!s) throw Error('Select the affected range to prepare the proposal.');
      try {message.proposal=await makeManual(s,routedProposal||result.intent,result.option);}
      catch(error) {throw Error('Could not prepare the proposal. '+error.message);}
    } else if (result.finding>=0) {
      if (!findings[result.finding]?.action) throw Error('The recommended workbook action is no longer available.');
      message.proposal=findings[result.finding];
    }
    if (result.intent==='audit') {
      state.workspaceMode='review';
      await doScan();
      message.text='Workbook review complete. Open the findings below for the affected cells and evidence.';
    }
    state.messages.push(message);
    return '';
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
    'Sending a question shares your prompt, up to 12 rows by 8 columns from the selection, up to 12 audit summaries, sampled sheet and column profiles, and recent chat via your Forma host and configured AI provider. If the assistant needs more evidence, it can request a bounded range or formula sample. Totals, counts and duplicates are computed locally. Workbook scans and writes run inside Excel. Session settings and undo data clear when the pane closes.')));
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
