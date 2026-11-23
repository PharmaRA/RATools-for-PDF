// One entry point for local qualification and the Windows portal build job.
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';
import { website } from './site-config.mjs';

const guard = pathToFileURL(path.join(website, 'scripts/recycle-guard.mjs')).href;
const env = {
  ...process.env,
  DOCS_SITE: 'https://ratools.qzz.io',
  DOCS_BASE: '/pdf',
  ASTRO_TELEMETRY_DISABLED: '1',
  NODE_OPTIONS: `${process.env.NODE_OPTIONS || ''} --import=${guard}`.trim(),
};
const steps = [
  ['Astro check', ['--import', guard, 'scripts/run-astro.mjs', 'check']],
  ['Build /pdf/', ['--import', guard, 'scripts/run-astro.mjs', 'build']],
  ['Check rendered site', ['scripts/check-site.mjs']],
];
for (const [label, args] of steps) {
  console.log(`[portal] ${label}`);
  const result = spawnSync(process.execPath, args, { cwd: website, env, stdio: 'inherit', windowsHide: true });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}
console.log('[portal] Ready: https://ratools.qzz.io/pdf/ — upload the contents of website/dist to the portal pdf/ directory.');
