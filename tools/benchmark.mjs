import {performance} from 'node:perf_hooks';
import {aggregationEvidence} from '../public/evidence.mjs';
import {workbookCase} from '../tests/fixtures/workbook-corpus.mjs';

for (const rows of [10,100,1000,10000,50000,100000]) {
  const {dataset,headers}=workbookCase(440,0,rows);
  const before=process.memoryUsage().heapUsed,t=performance.now();
  try {
    const result=aggregationEvidence(dataset,{dimension:headers[1],measure:headers[2]});
    console.log(JSON.stringify({rows,status:'processed',included:result.counts.included,ms:Math.round((performance.now()-t)*100)/100,
      heapDeltaMiB:Math.round((process.memoryUsage().heapUsed-before)/1048576*100)/100}));
  } catch {
    console.log(JSON.stringify({rows,status:'unsupported',reason:'selection cell or row budget'}));
  }
}
