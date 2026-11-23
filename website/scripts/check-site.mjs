import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { website } from './site-config.mjs';

const candidates = process.platform === 'win32'
  ? [['py', ['-3']], ['python', []]]
  : [['python3', []], ['python', []]];
for (const [command, prefix] of candidates) {
  const probe = spawnSync(command, [...prefix, '-c', 'import sys; sys.exit(sys.version_info < (3, 9))'], {
    windowsHide: true, stdio: 'ignore', timeout: 10000,
  });
  if (probe.error || probe.status !== 0) continue;
  const result = spawnSync(command, [...prefix, '-X', 'utf8', '-B', path.join(website, 'scripts/check-site.py'), ...process.argv.slice(2)], {
    windowsHide: true, stdio: 'inherit', cwd: website,
  });
  if (result.error) throw result.error;
  process.exit(result.status ?? 1);
}
console.error('Python 3.9+ is required for check:site; building and previewing only require Node.js.');
process.exit(1);
