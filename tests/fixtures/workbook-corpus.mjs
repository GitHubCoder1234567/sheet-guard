const families=[
  'Sales','Transactions','Customers','Inventory','Financial statements','Budget',
  'Employees','Education','Projects','Order fulfillment','Multi currency','Cross sheet',
  'Headerless','Multiple tables','Formula heavy','Messy records','Sparse table',
  'CSV import','Filtered records','Protected workbook','Mixed language','Nonstandard names',
  'Reordered columns','Repeated headers','Unusual formats',
];

function random(seed) {
  let state=seed>>>0;
  return n=>{state=(Math.imul(state,1664525)+1013904223)>>>0;return (state>>>8)%n;};
}

export function workbookCase(seed, familyIndex, rows=12) {
  const pick=random(seed), name=families[familyIndex];
  if (!name) throw Error('Unknown synthetic workbook family.');
  const headers=familyIndex===20?['Clave','Grupo','Montant']:['Record ID','Segment','Amount'];
  const derived=Array.from({length:rows},(_,i)=>({
    key:'S'+seed+'-'+i,group:['North','South','East','West'][pick(4)],amount:pick(201)-100,
  }));
  const expected=Object.create(null);
  // The oracle consumes the originating integer records, independent of worksheet shape and Forma.
  for (const row of derived) expected[row.group]=(expected[row.group]??0)+row.amount;
  const records=derived.map(r=>[r.key,r.group,r.amount]);
  let grid=[headers,...records], expectation='inferred', rawFormats;
  if (familyIndex===10) {
    rawFormats=grid.map((_,i)=>['General','General',i%2?'AUD 0':'USD 0']);
    expectation='mixed_currency';
  }
  if (familyIndex===12) {grid=records;expectation='ambiguous';}
  if (familyIndex===13) grid=grid.map(row=>[...row,'',...row]);
  if (familyIndex===15) {grid=[['Unverified import','',''],[],...grid].map(r=>[r[0]??'',r[1]??'',r[2]??'']);}
  if (familyIndex===16) grid=[...grid.slice(0,5),['','',''],...grid.slice(5)];
  if (familyIndex===22) grid=grid.map(row=>[row[2],row[0],row[1]]);
  if (familyIndex===23) {grid.splice(Math.min(5,grid.length),0,headers.slice());expectation='ambiguous';}
  if (familyIndex===24) rawFormats=grid.map(()=>['General','General','0.00;[Red](0.00)']);
  if ([13,16].includes(familyIndex)) expectation='multiple';
  const letter=String.fromCharCode(64+grid[0].length);
  const sheet=familyIndex===21?"Ops '东 😀":name+' '+seed;
  const formulas=familyIndex===14?grid.map((row,i)=>[row[0],row[1],i?'=SUM(0,'+row[2]+')':row[2]]):null;
  return {dataset:{sheet,sheetId:'synthetic-'+seed,address:'A1:'+letter+grid.length,values:grid,
    errors:grid.map(()=>[]),formats:rawFormats,formulas},expected,expectation,headers,
    objects:{protected:familyIndex===19,filteredRows:familyIndex===18?[4,6]:[],
      tables:familyIndex===13?['LeftTable','RightTable']:[],
      relatedSheets:familyIndex===11?['Lookup '+seed]:[]}};
}

export const corpus=()=>families.map((_,i)=>workbookCase(1000+i,i));
export const holdout=()=>families.map((_,i)=>workbookCase(83900+i*13,i,15+i%8));
