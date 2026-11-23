import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { syncDocs } from './sync-docs.mjs';
import { website, repository } from './site-config.mjs';

if (!globalThis.__ratoolsRecycleGuard) throw new Error('Use the guarded npm scripts.');
const command = process.argv[2];
if (!['dev', 'build', 'preview', 'check'].includes(command)) throw new Error('Unsupported documentation command.');
if (command !== 'preview') syncDocs();
if (command === 'build') {
  const output = path.join(website, 'dist');
  if (fs.existsSync(output)) fs.rmSync(output, { recursive: true }); // intercepted: Recycle Bin only
}
if (command === 'preview' && !fs.existsSync(path.join(website, 'dist/index.html'))) {
  throw new Error('Build the documentation first with npm run build.');
}
if (command === 'dev') {
  let pending;
  const watcher = fs.watch(path.join(repository, 'docs/manual'), () => {
    clearTimeout(pending);
    pending = setTimeout(() => {
      try { syncDocs({ quiet: true }); }
      catch (error) { console.error(`[manual sync] ${error.message}`); }
    }, 150);
  });
  process.once('exit', () => watcher.close());
}
process.env.ASTRO_TELEMETRY_DISABLED = '1';
const astroPackage = path.join(website, 'node_modules/astro');
const metadata = JSON.parse(fs.readFileSync(path.join(astroPackage, 'package.json'), 'utf8'));
const cli = path.resolve(astroPackage, metadata.bin.astro);
process.argv = [process.execPath, cli, command, ...process.argv.slice(3)];
await import(pathToFileURL(cli).href);
