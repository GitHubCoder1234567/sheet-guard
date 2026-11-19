export function plot(canvas, data, action) {
  const W = 440, H = 260;
  canvas.width = W; canvas.height = H;
  const g = canvas.getContext('2d');
  const p = data.points;
  const ink = action.color || '#14664a';
  const xLabel = action.xLabel ?? data.x, yLabel = action.yLabel ?? data.y;
  g.font = '12px system-ui';
  g.fillStyle = '#153128';
  g.textAlign = 'center';
  g.fillText(action.title,W/2,18,W-20);
  if (['pie','doughnut'].includes(action.type)) {
    const total = p.reduce((sum,[,v])=>sum+v,0);
    const colors = ['#14664a','#d78945','#41679d','#aa4e68','#539983','#7a548c','#65713a','#8a6851','#326575','#926522'];
    let start = -Math.PI/2;
    p.forEach(([label,value],i)=>{
      const stop = start + Math.PI*2*(value/total);
      g.beginPath(); g.moveTo(155,132); g.arc(155,132,86,start,stop); g.closePath();
      g.fillStyle = colors[i]; g.fill();
      if (action.legend) {
        g.fillRect(275,37+i*19,10,10);
        g.fillStyle = '#153128'; g.textAlign = 'left'; g.fillText(label,291,46+i*19,135);
      }
      start = stop;
    });
    if (action.type === 'doughnut') {
      g.beginPath(); g.arc(155,132,47,0,Math.PI*2); g.fillStyle='white'; g.fill();
    }
    return;
  }
  const left = 65, top = 43, width = 349, height = 148;
  const horizontal = action.type === 'bar';
  const scatter = action.type === 'scatter';
  const min = Math.min(0,...p.map(([,y])=>y));
  const max = Math.max(0,...p.map(([,y])=>y));
  const fraction = y => (y-min)/(max-min || 1);
  const bottom = top+height;
  const x0 = left + width*fraction(0), y0 = bottom-height*fraction(0);
  const lowX = scatter ? Math.min(...p.map(([, ,x])=>x)) : 0;
  const highX = scatter ? Math.max(...p.map(([, ,x])=>x)) : 0;
  const tick = v => Intl.NumberFormat(undefined,{notation:'compact',maximumFractionDigits:2}).format(v);
  g.font = '10px system-ui';
  for (let i = 0; i <= 4; i++) {
    g.strokeStyle = '#d6e2db'; g.beginPath();
    const y = bottom-height*i/4, x = left+width*i/4;
    if (horizontal) {g.moveTo(x,top);g.lineTo(x,bottom);} else {g.moveTo(left,y);g.lineTo(left+width,y);}
    g.stroke(); g.fillStyle = '#405850'; g.textAlign = horizontal ? 'center' : 'right';
    g.fillText(tick(min+(max-min)*i/4),horizontal?x:left-5,horizontal?bottom+14:y+3);
    if (scatter) {g.textAlign='center';g.fillText(tick(lowX+(highX-lowX)*i/4),x,bottom+14);}
  }
  const points = p.map(([,v,x],i) => [
    scatter ? left+width*((x-lowX)/(highX-lowX || 1)) : left+width*(i+.5)/p.length,
    bottom-height*fraction(v),
  ]);
  g.fillStyle=ink; g.strokeStyle=ink; g.lineWidth=2;
  if (action.type === 'line' || action.type === 'area') {
    g.beginPath(); points.forEach(([x,y],i)=>i ? g.lineTo(x,y) : g.moveTo(x,y));
    g.stroke();
    if (action.type === 'area') {
      g.lineTo(points.at(-1)[0],y0);g.lineTo(points[0][0],y0);g.closePath();
      g.globalAlpha=.28;g.fill();g.globalAlpha=1;
    }
  } else p.forEach(([,v],i)=>{
    const [x,y]=points[i];
    if (scatter) {g.beginPath();g.arc(x,y,3,0,Math.PI*2);g.fill();}
    else if (horizontal) {
      const end=left+width*fraction(v), band=height/p.length;
      g.fillRect(Math.min(x0,end),top+band*(i+.15),Math.abs(end-x0),band*.7);
    } else {
      const band=width/p.length;
      g.fillRect(x-band*.33,Math.min(y,y0),band*.66,Math.abs(y-y0));
    }
  });
  g.fillStyle='#153128';
  if (!scatter) p.forEach(([label],i)=>{
    if (i%Math.ceil(p.length/(horizontal?9:6))) return;
    g.textAlign=horizontal?'right':'center';
    g.fillText(label,horizontal?left-5:points[i][0],horizontal?top+height*(i+.5)/p.length+3:bottom+14,horizontal?55:54);
  });
  g.textAlign='center';
  g.fillText(horizontal?yLabel:xLabel,left+width/2,bottom+31,330);
  if (!horizontal) g.fillText(yLabel,left,top-6,120);
  else g.fillText(xLabel,32,top-6,60);
  if (action.legend) {g.fillStyle=ink;g.fillRect(left,237,9,9);g.textAlign='left';g.fillText(yLabel,left+14,245,300);}
}
