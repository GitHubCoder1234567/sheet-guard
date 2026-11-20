import {readFile} from 'node:fs/promises';
import {pathToFileURL} from 'node:url';

export const scenarios = [
  'connect', 'worksheet-resolution', 'renamed-sheet', 'deleted-sheet', 'table-and-totals',
  'named-range', 'ambiguous-data', 'typed-values', 'sales-ranking', 'column-calculation', 'composed-statistics',
  'source-selection', 'formula-errors', 'provider-probe', 'provider-fallback', 'report-privacy',
  'preview-cancel', 'stale-preview', 'apply-and-undo', 'protected-sheet', 'large-range',
  'discovery', 'resource-planning', 'formula-inspection', 'cross-sheet', 'key-relationships', 'ai-conversation', 'evidence-followup', 'conditional-ranking',
  'quality-audit', 'recovery', 'unsupported-api',
];
const hosts = ['Excel Web', 'Excel Windows', 'Excel Mac'];
const hasText = value => typeof value === 'string' && value.trim().length > 0;

export function verifyLiveReports(build,reports) {
  if (!/^[a-f0-9]{12}$/.test(build ?? '')) return ['Provide the candidate’s 12-character build ID.'];
  const problems=[];
  for (const host of hosts) {
    const matches=reports.filter(report=>report?.diagnostics?.host===host);
    if (matches.length!==1) {
      problems.push(host+' requires one evidence report.');
      continue;
    }
    const {diagnostics,tester,results}=matches[0];
    const date=Date.parse(diagnostics.date);
    if (diagnostics.schema!==1 || diagnostics.build!==build ||
        !Number.isFinite(date) || date>Date.now()+300000 || !hasText(diagnostics.reference) ||
        !hasText(tester) || !diagnostics.checks?.some(check=>check.category==='cell_readable'&&check.status==='passed')) {
      problems.push(host+' has incomplete diagnostics or incorrect build metadata.');
    }
    for (const id of scenarios) {
      const cases=Array.isArray(results)?results.filter(result=>result?.id===id):[];
      if (cases.length!==1 || cases[0].status!=='passed' || !hasText(cases[0].expected) || !hasText(cases[0].observed))
        problems.push(host+' requires a passing observation for '+id+'.');
    }
  }
  return problems;
}

if (process.argv[1] && import.meta.url===pathToFileURL(process.argv[1]).href) {
  try {
    const [, , build, ...files]=process.argv;
    const reports=await Promise.all(files.map(async file=>JSON.parse(await readFile(file,'utf8'))));
    const problems=verifyLiveReports(build,reports);
    if (problems.length) throw Error(problems.join('\n'));
    console.log('Live-host evidence recorded on all required hosts for '+build+'.');
  } catch(error) {
    console.error(error.message);
    process.exitCode=1;
  }
}
