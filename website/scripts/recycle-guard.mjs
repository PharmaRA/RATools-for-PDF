// Install before npm/Astro: every Node filesystem deletion is recycled or rejected.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { syncBuiltinESMExports } from 'node:module';

const website = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const allowed = ['dist', '.astro', 'node_modules', '.npm-cache'].map((part) => path.join(website, part));
const helper = path.join(website, 'scripts/recycle-path.ps1');

function inside(root, target) {
  const rel = path.relative(root, target);
  return !path.isAbsolute(rel) && rel !== '..' && !rel.startsWith(`..${path.sep}`);
}

function inspect(file) {
  const stat = fs.lstatSync(file);
  if (stat.isSymbolicLink() || (stat.isFile() && stat.nlink > 1)) {
    throw new Error('Cleanup rejected: linked filesystem entry requires review.');
  }
  if (stat.isDirectory()) {
    for (const name of fs.readdirSync(file)) inspect(path.join(file, name));
  }
}

function recycle(value) {
  const target = path.resolve(value instanceof URL ? fileURLToPath(value) : String(value));
  if (!allowed.some((root) => inside(root, target))) {
    throw new Error('Cleanup rejected outside website generated directories.');
  }
  let current = target;
  while (inside(website, current)) {
    if (fs.existsSync(current) && fs.lstatSync(current).isSymbolicLink()) {
      throw new Error('Cleanup rejected through a linked parent.');
    }
    if (current === website) break;
    current = path.dirname(current);
  }
  if (!fs.existsSync(target)) return;
  inspect(target);
  if (process.platform !== 'win32') throw new Error('Recycle-only cleanup currently requires Windows.');
  const result = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-File', helper, '-LiteralPath', target], {
    windowsHide: true, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'],
  });
  if (result.error || result.status !== 0) throw new Error('Recycle Bin cleanup failed; files retained.');
  console.log(`[recycle] ${path.relative(website, target)}`);
}

for (const name of ['rm', 'unlink', 'rmdir']) {
  fs[`${name}Sync`] = recycle;
  fs[name] = (file, ...args) => {
    const callback = args.at(-1);
    try { recycle(file); queueMicrotask(() => callback(null)); }
    catch (error) { queueMicrotask(() => callback(error)); }
  };
  fs.promises[name] = async (file) => recycle(file);
}
syncBuiltinESMExports();
globalThis.__ratoolsRecycleGuard = true;
