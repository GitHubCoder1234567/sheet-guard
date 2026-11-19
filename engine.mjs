export const LIMIT = 2400;
export const chartTypes = {
  line: 'Line', bar: 'BarClustered', column: 'ColumnClustered', area: 'Area',
  pie: 'Pie', doughnut: 'Doughnut', scatter: 'XYScatter',
};
export const palettes = {
  Professional: '#14664a', Financial: '#203b57', Corporate: '#343575',
  Minimal: '#454c49', Modern: '#00746d', Custom: '#385c81',
  Finance:'#203b57', 'Executive Report':'#284354', 'Data Analysis':'#245666',
};
export const categories = ['Errors', 'Data Quality', 'Formatting', 'Visualisations', 'Analysis', 'Automation'];
export const col = i => {
  let name = '';
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) name = String.fromCharCode(65 + (i - 1) % 26) + name;
  return name;
};
export function a1(row, column, rows = 1, columns = 1) {
  const start = col(column) + (row + 1);
  return rows === 1 && columns === 1 ? start : start + ':' + col(column + columns - 1) + (row + rows);
}
export function rect(address) {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})(?::([A-Z]{1,3})([1-9]\d{0,6}))?$/.exec(address);
  if (!match) throw Error('Use a local A1 range, such as A1:B12.');
  const [, x, y, xx = x, yy = y] = match;
  const index = name => [...name].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const r = +y - 1, c = index(x), rows = +yy - r, columns = index(xx) - c + 1;
  if (rows < 1 || columns < 1 || r + rows > 1048576 || c + columns > 16384 || rows * columns > LIMIT)
    throw Error('Range exceeds the supported size or Excel bounds.');
  return {r, c, rows, columns};
}
const blank = v => v === '' || v === null;
const numeric = v => typeof v === 'number' && Number.isFinite(v);
export const trimText = v => typeof v==='string' && /[A-Za-z]/.test(v.trim()) &&
  /^[A-Za-z]/.test(v.trim()) && v.trim()!==v ? v.trim() : null;
const text = v => typeof v === 'string' && !!v.trim();
export const formula = (s, r, c) => typeof s.formulas[r][c] === 'string' &&
  s.formulas[r][c].startsWith('=') && s.values[r][c] !== s.formulas[r][c];
// Only single-cell relative arithmetic between two same-row input cells.
export function correction(expression, column) {
  const m = /^=RC\[(-?\d{1,5})\]([+*\/-])RC\[(-?\d{1,5})\]$/.exec(expression);
  return !!m && [Number(m[1]), Number(m[3])].every(delta => delta !== 0 && column + delta >= 0 && column + delta < 16384);
}
export function chartData(s, type) {
  const [header, ...body] = s.values;
  if (header.length !== 2 || !header.every(text) || body.length < 3 || body.length > 80 ||
      body.some(([x,y]) => !(text(x) || numeric(x)) || !numeric(y)) ||
      s.types.some(row => row.includes('Error'))) throw Error('Charts require two headers and 3–80 complete data rows with numeric measures.');
  if (type === 'scatter' && !body.every(([x]) => numeric(x))) throw Error('Scatter charts require two numeric columns.');
  if (['pie', 'doughnut'].includes(type) && (body.length > 10 || body.some(([,y]) => y < 0) ||
      !body.some(([,y]) => y > 0) || new Set(body.map(([x]) => x)).size !== body.length))
    throw Error('Composition charts require up to ten distinct categories and nonnegative values with a positive total.');
  return {x: header[0], y: header[1], points: body.map(([x,y], i) => [s.display?.[i+1]?.[0] ?? String(x), y, x])};
}
function intent(s) {
  const header = String(s.values[0][0]).trim();
  if (/^(date|month|week|quarter|year|period)$/i.test(header)) {
    if (!s.values.slice(1).every(([x]) => numeric(x) || /^(?:\d{4}(?:-\d{2}(?:-\d{2})?)?|jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)$/i.test(String(x).trim()))) return null;
    return 'line';
  }
  if (s.values.slice(1).every(([x]) => numeric(x))) return 'scatter';
  if (s.values.slice(1).every(([x]) => text(x)) && new Set(s.values.slice(1).map(([x]) => x.trim().toLowerCase())).size === s.values.length - 1) return 'bar';
  return null;
}
function invalidDate(v) {
  if (typeof v !== 'string') return false;
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v.trim());
  const slash = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(v.trim());
  const valid = (y,m,d) => y >= 1900 && m >= 1 && m <= 12 && d >= 1 &&
    d <= new Date(Date.UTC(y,m,0)).getUTCDate();
  if (iso) return !valid(+iso[1],+iso[2],+iso[3]);
  return !!slash && !valid(+slash[3],+slash[1],+slash[2]) && !valid(+slash[3],+slash[2],+slash[1]);
}
function localCycles(s) {
  const nodes = new Map();
  const token = '\\$?[A-Z]{1,3}\\$?[1-9]\\d{0,6}';
  const arithmetic = new RegExp('^='+token+'(?:[+*\\/-]'+token+')*$','i');
  for (let r=0;r<s.values.length;r++) for (let c=0;c<s.values[r].length;c++) {
    const expression=s.formulas[r][c];
    if (formula(s,r,c) && arithmetic.test(expression)) nodes.set(a1(s.row+r,s.column+c),
      expression.toUpperCase().match(new RegExp(token,'g')).map(x=>x.replaceAll('$','')));
  }
  const finished=new Set(), path=new Set(), found=new Set();
  // Depth-first traversal with an explicit stack for large ranges.
  for (const start of nodes.keys()) {
    if (finished.has(start)) continue;
    const stack=[[start,0]];
    path.add(start);
    while (stack.length) {
      const frame=stack.at(-1), [cell,i]=frame, edges=nodes.get(cell);
      if (i===edges.length) {path.delete(cell);finished.add(cell);stack.pop();continue;}
      frame[1]++;
      const next=edges[i];
      if (path.has(next)) {
        for (const [member] of stack.slice(stack.findIndex(([x])=>x===next))) found.add(member);
      } else if (nodes.has(next) && !finished.has(next)) {path.add(next);stack.push([next,0]);}
    }
  }
  return [...found];
}
// This lexer reports A1 references for inspection. It does not rewrite formulas.
export function formulaReferences(expression) {
  const opaque=/"(?:[^"]|"")*"|\[(?:[^\[\]]|\[[^\]]*\])*\]/g;
  const masked=expression.replace(opaque,m=>' '.repeat(m.length));
  const refs=/(?<![.\w])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(\$?[A-Z]{1,3}\$?[1-9]\d{0,6})(?::(\$?[A-Z]{1,3}\$?[1-9]\d{0,6}))?(?![.\w(])/gi;
  return [...masked.matchAll(refs)].flatMap(m=>{
    // An external book reference and a 3-D range are outside this parser.
    if (/\]\s*$|:\s*$/.test(expression.slice(0,m.index))) return [];
    const first=m[3].replaceAll('$','').toUpperCase(), last=m[4]?.replaceAll('$','').toUpperCase();
    try {rect(first);if (last) rect(last);}
    catch {return [];}
    return [{sheet:m[1]?.replaceAll("''","'")??m[2],first,last,absolute:m[3].includes('$') || !!m[4]?.includes('$')}];
  });
}

export function profile(s) {
  const columns=s.values[0]?.map((label,c)=>{
    const counts={numbers:0,text:0,blank:0,formulas:0,errors:0};
    for (let r=1;r<s.values.length;r++) {
      const v=s.values[r][c];
      if (s.types[r][c]==='Error') counts.errors++;
      else if (blank(v)) counts.blank++;
      else if (numeric(v)) counts.numbers++;
      else if (text(v)) counts.text++;
      if (formula(s,r,c)) counts.formulas++;
    }
    const format=s.formats.slice(1).map(row=>row[c]).find(v=>v&&v!=='General')||'';
    const type=/\bid\b|identifier|reference number/i.test(String(label))?'Identifier':
      /\bdate\b|^month$|^year$/i.test(String(label))?'Date or period':
      format.includes('%')?'Percentage':/AUD|USD|EUR|GBP|[$€£]/.test(format)?'Monetary':
      counts.numbers>counts.text?'Numeric':counts.text?'Text':'Other';
    return {label:String(label||col(s.column+c)),type,...counts};
  })??[];
  const purpose=/(sales|revenue|orders)/i.test(columns.map(c=>c.label).join(' '))?'Possible sales reporting':
    /(inventory|stock|sku)/i.test(columns.map(c=>c.label).join(' '))?'Possible inventory tracking':
    /(budget|actual|variance)/i.test(columns.map(c=>c.label).join(' '))?'Possible budget reporting':null;
  return {columns,purpose};
}

function dependencyEvidence(s,sheets) {
  const dependencies=new Map(), broken=[];
  const lookup=new Map(sheets.map(sheet=>[sheet.name.toLowerCase(),sheet]));
  for (let r=0;r<s.values.length;r++) for (let c=0;c<s.values[r].length;c++) {
    if (!formula(s,r,c)) continue;
    const expression=s.formulas[r][c];
    const address=a1(s.row+r,s.column+c);
    if (/#REF!/i.test(expression.replace(/"(?:[^"]|"")*"/g,''))) broken.push(address);
    for (const reference of formulaReferences(expression.slice(0,4096)).slice(0,64)) {
      if (!reference.sheet || reference.sheet.toLowerCase()===s.name.toLowerCase()) continue;
      const key=reference.sheet.toLowerCase();
      if (!dependencies.has(key)) dependencies.set(key,{name:reference.sheet,count:0,example:address,error:0});
      const item=dependencies.get(key);
      item.count++;
      const other=lookup.get(key);
      if (other && !reference.last) {
        const {r:row,c:column}=rect(reference.first);
        if (other.types[row-other.row]?.[column-other.column]==='Error') item.error++;
      }
    }
  }
  return {dependencies:[...dependencies.values()],broken};
}

export const additiveMeasure = (label,action={}) =>
  /^(?:total )?(?:sales|revenue|units|quantity|orders)(?: \([A-Z]{3}\))?$/i.test(label) ||
  action.confirmedMetric === label;

export function audit(sheets) {
  const findings = [];
  for (const s of sheets) {
    if (!s.values.length) continue;
    const add = (category, priority, title, detail, address, action, evidence = 'Observed') =>
      findings.push({category, priority, title, detail, sheet: s.name, sheetId: s.id, address, action, evidence,
        id: JSON.stringify([s.id, category, title, address])});
    const rows = s.values.length, widths = s.values[0].length;
    const addr = (r,c,rr=1,cc=1) => a1(s.row+r,s.column+c,rr,cc);
    const errors = [], spaces = [], trimCandidates=[], duplicate = [], seen = new Set();
    for (let r = 0; r < rows; r++) {
      const row = s.values[r];
      if (r && row.some(v => !blank(v))) {
        const key = JSON.stringify(row);
        if (seen.has(key)) duplicate.push(addr(r,0,1,widths));
        seen.add(key);
      }
      row.forEach((v,c) => {
        if (s.types[r][c] === 'Error') errors.push(addr(r,c) + ' ' + v);
        if (text(v) && v.trim() !== v && !formula(s,r,c)) {
          spaces.push(addr(r,c));
          if (s.types[r][c]==='String' && trimText(v)!==null) trimCandidates.push(addr(r,c));
        }
      });
    }
    const grouped = (category, priority, title, cells, detail) => {
      if (!cells.length) return;
      add(category, priority, title, detail + ' ' + cells.slice(0, 8).join(', ') +
        (cells.length > 8 ? ' … (' + cells.length + ' total)' : ''), addr(0,0,rows,widths));
      findings.at(-1).focusAddress=cells[0].split(' ')[0];
    };
    const {dependencies,broken}=dependencyEvidence(s,sheets);
    grouped('Errors', 97, 'Broken formula references', broken, 'These formula expressions contain #REF!. Excel lost a referenced cell or range.');
    for (const dependency of dependencies.slice(0,6)) add(
      dependency.error?'Errors':'Analysis',dependency.error?86:29,
      dependency.error?'Referenced sheet contains errors · '+dependency.name:'Cross-sheet dependency · '+dependency.name,
      dependency.count+' direct reference(s) link here. '+(dependency.error?
        dependency.error+' referenced cell(s) returned an error in the sampled data.':'Changes to that sheet may affect these calculations.'),
      dependency.example,undefined,'Parsed A1 references');
    grouped('Errors', 99, 'Circular formula references', localCycles(s), 'Local arithmetic references form a cycle.');
    grouped('Errors', 100, 'Excel error values', errors, 'Inspect these calculation errors.');
    grouped('Data Quality', 90, 'Duplicate records', duplicate, 'Confirm which records to retain.');
    grouped('Data Quality', 63, 'Leading or trailing whitespace', spaces, 'Review these text cells before standardising them.');
    for (const address of trimCandidates.slice(0,3))
      add('Data Quality',61,'Trim text whitespace','Remove leading and trailing spaces from this text cell.',address,
        {kind:'trim',sheetId:s.id,address},'Literal text');
    for (let c = 0; c < widths; c++) {
      const invalidDates = [], formats = new Map(), numbersAsText = [], missing = [], categoryVariants = new Map(), sample = [];
      const label = String(s.values[0][c] || col(s.column+c));
      const patterns=new Map();
      for (let r = 1; r < rows; r++) {
        const v = s.values[r][c];
        if (/\bdate\b/i.test(label) && invalidDate(v)) invalidDates.push(addr(r,c));
        if (numeric(v)) {
          const fmt = s.formats[r][c];
          formats.set(fmt, (formats.get(fmt) || 0) + 1);
          sample.push([v,r]);
        }
        if (text(v) && /^[-+]?\d+(?:\.\d+)?$/.test(v.trim())) numbersAsText.push(r);
        if (blank(v) && s.values[r].some(x => !blank(x))) missing.push(addr(r,c));
        if (text(v) && !formula(s,r,c)) {
          const key = v.trim().toLowerCase();
          if (!categoryVariants.has(key)) categoryVariants.set(key, new Set());
          categoryVariants.get(key).add(v);
        }
        if (formula(s,r,c) && typeof s.r1c1[r][c]==='string') {
          const pattern=s.r1c1[r][c];
          if (!patterns.has(pattern)) patterns.set(pattern,[]);
          patterns.get(pattern).push(r);
        }
        if (r < 3 || r+2 >= rows) continue;
        const neighbours = [-2,-1,1,2].map(d => s.r1c1[r+d][c]);
        const expected = neighbours[0];
        if (![-2,-1,1,2].every(d => formula(s,r+d,c)) || !correction(expected,s.column+c) ||
            !neighbours.every(f => f === expected) || s.r1c1[r][c] === expected || !s.values[r].some((x,i) => i !== c && !blank(x))) continue;
        const address = addr(r,c);
        add('Errors', 80, formula(s,r,c) ? 'Formula differs from nearby rows' : 'Calculated column has an exception',
          'Four adjacent rows use ' + expected + '. Review this exception before applying the pattern.',
          address, {kind:'formula', sheetId:s.id, address, value:expected}, 'Adjacent formula pattern');
      }
      if (patterns.size>1) {
        const ordered=[...patterns].sort((a,b)=>b[1].length-a[1].length);
        const [dominant,matches]=ordered[0], formulaCount=ordered.reduce((n,[,items])=>n+items.length,0);
        if (matches.length>=5 && matches.length/formulaCount>=0.75) {
          for (const [pattern,exceptions] of ordered.slice(1)) {
            if (exceptions.length!==1 || correction(dominant,s.column+c) || correction(pattern,s.column+c)) continue;
            const r=exceptions[0];
            if (r<=matches[0] || r>=matches.at(-1)) continue;
            add('Errors',71,'Formula pattern exception · '+label,
              matches.length+' sampled formulas share the pattern '+dominant.slice(0,180)+
              '. This cell uses '+pattern.slice(0,180)+'. Confirm whether it represents a deliberate exception.',
              addr(r,c),undefined,'Dominant column pattern');
          }
        }
      }
      if (formats.size > 1) {
        const address=addr(1,c,rows-1);
        const expected=/\bdate\b/i.test(label)?'yyyy-mm-dd':
          /\b(?:discount|margin|percentage|percent)\b/i.test(label)?'0.0%':null;
        const canStandardise=expected && sample.length>=5 && (formats.get(expected)||0)>=.8*sample.length;
        add('Formatting', 55, 'Mixed number formats · ' + label,
          'There are ' + formats.size + ' number formats on numeric cells. Review their units before standardising.',address,
          canStandardise?{kind:'number',sheetId:s.id,address,value:expected}:undefined,
          canStandardise?'Dominant display format':'Observed');
      }
      if (sample.length && numbersAsText.length) grouped('Data Quality', 75, 'Numeric text · ' + label, numbersAsText.map(r => addr(r,c)), 'These may be numbers stored as text.');
      grouped('Data Quality', 72, 'Invalid dates · ' + label, invalidDates, 'These date strings are impossible in the recognised date formats.');
      grouped('Data Quality', 65, 'Missing entries · ' + label, missing, 'These records have a blank field.');
      const variants = [...categoryVariants.values()].filter(v => v.size > 1);
      if (variants.length) add('Data Quality', 62, 'Text variants · ' + label,
        variants.slice(0,4).map(v=>[...v].map(x=>JSON.stringify(x)).join(' / ')).join(', '), addr(1,c,rows-1));
      if (sample.length >= 8) {
        sample.sort((a,b)=>a[0]-b[0]);
        const q1=sample[Math.floor((sample.length-1)*.25)][0], q3=sample[Math.ceil((sample.length-1)*.75)][0];
        const iqr=q3-q1;
        if (iqr > 0) grouped('Data Quality', 45, 'Numerical outliers · '+label,
          sample.filter(([v])=>v<q1-1.5*iqr || v>q3+1.5*iqr).map(([,r])=>addr(r,c)), 'Outside the 1.5× interquartile fences.');
      }
    }
    if (widths === 2 && !s.partial && s.chartCount === 0) {
      const type = intent(s);
      if (!type) continue;
      let data;
      try { data = chartData(s,type); } catch { continue; }
      const address = addr(0,0,rows,widths);
      add('Visualisations', 40, type === 'line' ? 'Show the trend over time' : type === 'bar' ? 'Compare categories' : 'Compare numeric variables',
        data.y + ' by ' + data.x, address, {kind:'chart', sheetId:s.id, address, type,
          title:data.y + ' by ' + data.x, legend:false, xLabel:data.x, yLabel:data.y}, 'Complete two-column dataset');
      if (type !== 'scatter' && additiveMeasure(data.y))
        add('Analysis', 35, 'Summarise ' + data.y, 'Create a separate worksheet with a linked total, observation count and chart.',
          address, {kind:'dashboard', sheetId:s.id, address, type, title:data.y + ' overview', legend:false}, 'Additive measure inferred from heading');
      else if (type!=='scatter' && !/rate|ratio|percent|margin|average|%/i.test(data.y)) {
        add('Analysis',31,'Confirm how to summarise '+data.y,
          'A dashboard total depends on whether this metric can be added across rows.',address,undefined,'Business definition required');
        findings.at(-1).clarification={metric:data.y,type,title:data.y+' overview'};
      }
    }
  }
  return findings.sort((a,b) => b.priority-a.priority);
}
export function validate(action) {
  if (!action || typeof action.sheetId !== 'string' || !action.sheetId || typeof action.address !== 'string') throw Error('Missing workbook target.');
  const bounds = rect(action.address);
  if (action.kind === 'formula') {
    if (bounds.rows !== 1 || bounds.columns !== 1 || !correction(action.value,bounds.c)) throw Error('Only bounded single-cell arithmetic repairs are supported.');
  } else if (action.kind === 'trim') {
    if (bounds.rows!==1 || bounds.columns!==1) throw Error('Text cleanup is limited to one cell.');
  } else if (action.kind === 'number') {
    if (!['0.0%', '#,##0.00', 'yyyy-mm-dd', '0'].includes(action.value)) throw Error('Unsupported number format.');
  } else if (action.kind === 'style') {
    if (!Object.hasOwn(palettes,action.value)) throw Error('Unsupported style.');
    if (action.color !== undefined && !/^#[0-9a-f]{6}$/i.test(action.color)) throw Error('Invalid colour.');
  } else if (action.kind === 'heatmap' || action.kind === 'negative' || action.kind === 'dataBar') {
    if (bounds.columns !== 1) throw Error('Choose one numeric column.');
  } else if (action.kind === 'chart' || action.kind === 'dashboard') {
    if (bounds.columns !== 2 || bounds.rows < 4 || bounds.rows > 81 || !Object.hasOwn(chartTypes,action.type) ||
        typeof action.title !== 'string' || !action.title.trim() || action.title.length > 120 ||
        /^\s*[=+\-@]/.test(action.title) || typeof action.legend !== 'boolean') throw Error('Invalid chart configuration.');
    if (action.confirmedMetric !== undefined &&
        (action.kind !== 'dashboard' || typeof action.confirmedMetric !== 'string' ||
         !action.confirmedMetric.trim() || action.confirmedMetric.length > 120))
      throw Error('Invalid confirmed dashboard metric.');
    for (const label of [action.xLabel,action.yLabel]) {
      if (label !== undefined && (typeof label !== 'string' || label.length > 120)) throw Error('Invalid axis label.');
    }
    if (action.color !== undefined && !/^#[0-9a-f]{6}$/i.test(action.color)) throw Error('Invalid colour.');
    if (action.destination !== undefined && (action.kind !== 'chart' || typeof action.destination !== 'string' || !action.destination)) throw Error('Invalid chart destination.');
    if (action.position !== undefined) {
      const box = rect(action.position);
      if (action.kind !== 'chart' || box.rows < 4 || box.columns < 3) throw Error('Chart position needs a rectangle at least 4 rows by 3 columns.');
    }
  } else throw Error('Unsupported operation.');
  return bounds;
}
export function conflict(a,b) {
  if (a.sheetId !== b.sheetId) return false;
  const p = rect(a.address), q = rect(b.address);
  return p.r < q.r+q.rows && q.r < p.r+p.rows && p.c < q.c+q.columns && q.c < p.c+p.columns;
}

export function formulaPreview(plan) {
  const {action,snapshot,formulaEvidence}=plan;
  const match=/^=RC\[(-?\d+)\]([+*\/-])RC\[(-?\d+)\]$/.exec(action.value);
  if (!match || !formulaEvidence) return null;
  const [,left,operator,right]=match;
  const [x,y]=formulaEvidence.inputs;
  const result={'+':()=>x+y,'-':()=>x-y,'*':()=>x*y,'/':()=>x/y}[operator]();
  const band=formulaEvidence.band;
  return {
    before:snapshot.formulas[0][0],
    after:'='+a1(snapshot.row,snapshot.column+Number(left))+operator+a1(snapshot.row,snapshot.column+Number(right)),
    current:snapshot.display[0][0],
    estimate:Number.isFinite(result) && !(operator==='/' && y===0)?result:null,
    neighbours:band.formulas.flatMap(([cell],i)=>i===2?[]:[[a1(band.row+i,band.column),cell]]),
  };
}

export function formatted(value,format,fallback) {
  if (typeof value!=='number') return fallback;
  if (format==='yyyy-mm-dd') return null;
  const settings={
    '0.0%':{style:'percent',minimumFractionDigits:1,maximumFractionDigits:1},
    '#,##0.00':{minimumFractionDigits:2,maximumFractionDigits:2},
    '0':{maximumFractionDigits:0,useGrouping:false},
  }[format];
  return settings?new Intl.NumberFormat(undefined,settings).format(value):fallback;
}

export function previewRows(plan) {
  const {snapshot:s,action:a}=plan;
  const sample=s.values.slice(0,5).flatMap((row,r)=>row.slice(0,4).map((value,c)=>({
    address:a1(s.row+r,s.column+c),value,before:s.display[r][c],format:s.formats[r][c],
    after:a.kind==='number'?formatted(value,a.value,s.display[r][c]):a.kind==='trim'?trimText(value):s.display[r][c],
  })));
  return sample;
}
export function headingStyle(action) {
  const fill=(action.color || palettes[action.value]).toUpperCase();
  const rgb=fill.slice(1).match(/../g).map(hex=>{
    const value=parseInt(hex,16)/255;
    return value<=0.04045?value/12.92:((value+0.055)/1.055)**2.4;
  });
  const luminance=rgb[0]*0.2126+rgb[1]*0.7152+rgb[2]*0.0722;
  return [true,luminance>0.179?'#000000':'#FFFFFF',fill];
}
