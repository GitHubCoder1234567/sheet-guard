import {audit, categories, chartTypes, palettes, conflict, validate} from './engine.mjs';
import {scan, selection, preview, apply, rollback, supported} from './excel.mjs';
import {plot} from './plot.mjs';

const nav = document.querySelector('#nav'), main = document.querySelector('#main');
const notice = document.querySelector('#notice'), connection = document.querySelector('#connection');
const tabs = ['Chat','Audit','Recommendations','Visualisations','History','Settings'];
const state = {
  tab:'Chat', ready:false, busy:false, findings:[], snapshots:[], dismissed:new Set(),
  previews:new Map(), edits:new Map(), checked:new Set(), history:[], messages:[],
  filter:'All', consent:false, draft:'', scanned:false, manualKind:'chart', accessPassword:'',
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
async function doScan() {
  state.scanController=new AbortController();
  render();
  try {
    const {snapshots,omitted}=await scan(state.scanController.signal,(done,total)=>{
      notice.textContent='Scanned '+done+' populated worksheet(s), '+total+' visible.';
    });
    state.snapshots=snapshots;state.omitted=omitted;state.scanned=true;
    const manual=state.findings.filter(f=>f.manual);
    state.findings=[...audit(snapshots),...manual];
    state.previews.clear();state.checked.clear();
    return 'Audit complete.';
  } finally {state.scanController=null;}
}
function auditView() {
  const controls=node('div',undefined,'actions');
  controls.append(button('Scan visible sheets',()=>run(doScan,'Scanning workbook…'),{primary:true}));
  if (state.scanController) controls.append(button('Cancel scan',()=>state.scanController.abort(),{during:true}));
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
  main.append(stats);
  const partial=state.snapshots.filter(s=>s.partial).map(s=>s.name);
  if (partial.length || state.omitted) main.append(node('p',
    (partial.length?'Sampled ranges on '+partial.join(', ')+'. ':'')+
    (state.omitted?state.omitted+' visible sheet(s) beyond the scan limit.':''),'muted'));
  main.append(button('Review recommendations',()=>setTab('Recommendations')));
}

const labels={chart:'Create chart',dashboard:'Create dashboard',style:'Style headings',number:'Change number display',
  formula:'Repair formula',negative:'Highlight negatives',heatmap:'Apply a colour scale',dataBar:'Add data bars'};
const numberFormats={percent:'0.0%',decimal:'#,##0.00',date:'yyyy-mm-dd',integer:'0'};
function manualPanel(visual) {
  const card=block(node('h3','From the current selection'));
  const options=Object.entries(labels).filter(([kind])=>kind!=='formula' && (!visual || ['chart','dashboard','heatmap','dataBar'].includes(kind)));
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
    filter.addEventListener('change',()=>{state.filter=filter.value;render();});
    main.append(field('Category',filter));
  }
  const list=state.findings.filter(f=>!state.dismissed.has(f.id) && (visual?
    ['Visualisations','Analysis'].includes(f.category):state.filter==='All'||f.category===state.filter));
  const candidates=list.slice(0,32).filter(f=>state.checked.has(f.id)&&state.previews.has(f.id));
  if (candidates.length>1) {
    const bulk=button('Approve '+candidates.length+' selected changes',()=>run(async()=>{
    const plans=candidates.map(f=>state.previews.get(f.id));
    for (let i=0;i<plans.length;i++) for (let j=0;j<i;j++)
      if (conflict(plans[i].action,plans[j].action) || [plans[i],plans[j]].some(p=>p.action.kind==='chart'))
        throw Error('The batch contains overlapping sources or chart placement. Approve these changes individually.');
    let completed=0;
    try {for (const f of candidates) {await execute(f);completed++;}}
    catch (error) {throw Error(completed+' verified before the batch stopped. '+error.message);}
    return completed+' changes verified.';
  },'Checking and applying selected changes…'),{primary:true});
    bulk.dataset.bulk='true';main.append(bulk);
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
    state.history.unshift({title:f.title,result:'Verified · '+result.description,recovery:result.recovery,time:new Date()});
    // Keep rollback state in memory only and limit the number of retained ranges.
    state.history.slice(25).forEach(item=>delete item.recovery);
    state.dismissed.add(f.id);state.previews.delete(f.id);state.edits.delete(f.id);state.checked.delete(f.id);
  } catch (error) {
    state.history.unshift({title:f.title,result:'Unverified · '+error.message,time:new Date()});
    state.previews.delete(f.id);state.checked.delete(f.id);
    throw error;
  }
}
function card(f) {
  const card=block(node('div',f.category+' · '+f.evidence,'meta'),node('h3',f.title),
    node('div',findingAddress(f),'muted'),node('p',f.detail));
  const action=state.edits.get(f.id)?.action || f.action;
  const actions=node('div',undefined,'actions'), p=state.previews.get(f.id);
  if (action && available(action)) {
    actions.append(button(p?'Refresh preview':'Preview',()=>run(async()=>{
      state.previews.delete(f.id);state.checked.delete(f.id);
      state.previews.set(f.id,await preview(action));
      return 'Review the preview before approving.';
    },'Reading preview…')));
    if (['chart','dashboard','style'].includes(action.kind))
      actions.append(button(state.edits.has(f.id)?'Close customisation':'Customise',()=>{
        if (state.edits.has(f.id)) {state.edits.delete(f.id);state.previews.delete(f.id);state.checked.delete(f.id);render();}
        else run(async()=>{
          const sheets=action.kind==='chart'?await Excel.run(async ctx=>{
            const collection=ctx.workbook.worksheets;
            collection.load('items/id,items/name,items/visibility');await ctx.sync();
            return collection.items.filter(s=>s.visibility==='Visible').map(s=>[s.id,s.name]);
          }):[];
          state.edits.set(f.id,{action:{...action},sheets});
          return '';
        },'Loading options…');
      }));
    if (p) actions.append(button('Accept',()=>run(async()=>{await execute(f);return 'Change verified in Excel.';},'Applying approved change…'),{primary:true}));
  } else if (action) card.append(node('p','This operation requires a newer Excel API.','muted'));
  actions.append(button('Decline',()=>{
    state.dismissed.add(f.id);state.previews.delete(f.id);state.edits.delete(f.id);state.checked.delete(f.id);
    state.history.unshift({title:f.title,result:'Declined · '+findingAddress(f),time:new Date()});render();
  }));
  card.append(actions);
  if (state.edits.has(f.id)) card.append(customisation(f));
  if (p) {
    card.append(previewPanel(p));
    const checkbox=node('input');checkbox.type='checkbox';checkbox.checked=state.checked.has(f.id);
    checkbox.disabled=state.busy;
    checkbox.addEventListener('change',()=>{
      if (checkbox.checked) state.checked.add(f.id);else state.checked.delete(f.id);
      render();
    });
    const wrap=node('label',undefined,'check');wrap.append(checkbox,node('span','Select for bulk approval'));card.append(wrap);
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
          if (control.textContent==='Accept' || control.type==='checkbox'&&!details.contains(control)) control.disabled=true;
        }
        notice.textContent='Configuration changed. Refresh preview before approving.';
      }
      if (key==='type') render();
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
function previewPanel(p) {
  const section=node('section',undefined,'details'), {action:a,snapshot:s,data}=p;
  section.append(node('h3','Change preview'),node('p',s.name+'!'+a.address,'meta'));
  if (data) {
    section.append(node('p',a.type+' · '+a.title),node('p',(a.xLabel||data.x)+' / '+(a.yLabel||data.y),'muted'));
    if (a.kind==='dashboard') section.append(node('p',
      'New worksheet · Total '+data.points.reduce((n,[,v])=>n+v,0).toLocaleString()+' · '+data.points.length+' observations'));
    else section.append(node('p',p.destination+'!'+p.position,'muted'));
    const canvas=node('canvas');
    canvas.setAttribute('role','img');
    canvas.setAttribute('aria-label',a.type+' chart of '+data.y+' by '+data.x);
    section.append(canvas);plot(canvas,data,a);
    const table=node('details');
    table.append(node('summary','Chart data'),node('pre',data.points.map(([x,y])=>x+'  '+y).join('\n')));
    section.append(table,node('p','Excel controls the final chart rendering.','muted'));
  } else if (a.kind==='formula') {
    section.append(node('pre','Before  '+s.r1c1[0][0]+'\nAfter   '+a.value),node('p','Formula references are shown in R1C1 notation.','muted'));
  } else if (a.kind==='number') {
    section.append(node('pre','Current format · '+s.formats[0][0]+'\nProposed format · '+a.value),node('p','Changes display formatting for the entire range.','muted'));
  } else if (a.kind==='style') {
    const sample=node('p',s.display[0].slice(0,3).join('  ·  ') || 'Heading style');
    sample.style.background=a.color||palettes[a.value];sample.style.color='white';sample.style.fontWeight='bold';sample.style.padding='6px';
    section.append(sample,node('p','Apply '+a.value+' font and fill to the entire range.','muted'));
  } else section.append(node('p',labels[a.kind]+' for the selected numeric cells. Value and formula contents are preserved.'));
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
    const controller=new AbortController(), timeout=setTimeout(()=>controller.abort(),28000);
    let result;
    try {
      const headers={'Content-Type':'application/json'};
      if (state.accessPassword) {
        const bytes=new TextEncoder().encode('sheetguard:'+state.accessPassword);
        headers.Authorization='Basic '+btoa(String.fromCharCode(...bytes));
      }
      const response=await fetch('/api/agent',{method:'POST',headers,
        body:JSON.stringify({question,selection:context,findings:findings.map(f=>({title:f.title,detail:f.detail.slice(0,240),
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
    'Sending a question shares your prompt, up to 12 rows by 8 columns from the selection, up to 12 audit summaries and recent chat via your SheetGuard host and OpenAI. Workbook scans and writes run inside Excel. Session settings and undo data clear when the pane closes.')));
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
