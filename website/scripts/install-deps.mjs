import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';
import './recycle-guard.mjs';
import { website } from './site-config.mjs';

const candidates = [
  path.join(path.dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js'),
  path.join(path.dirname(process.execPath), '../lib/node_modules/npm/bin/npm-cli.js'),
];
const npm = candidates.find((candidate) => fs.existsSync(candidate));
if (!npm) throw new Error('npm CLI was not found beside Node.js. Install the standard Node.js distribution.');
const modules = path.join(website, 'node_modules');
if (fs.existsSync(modules)) fs.rmSync(modules, { recursive: true }); // guarded: recycle only
const guard = pathToFileURL(path.join(website, 'scripts/recycle-guard.mjs')).href;
const result = spawnSync(process.execPath, ['--import', guard, npm, 'ci', '--ignore-scripts', '--no-audit', '--no-fund',
  '--cache', path.join(website, '.npm-cache'), ...process.argv.slice(2)], {
  cwd: website,
  stdio: 'inherit',
  windowsHide: true,
  env: { ...process.env, ASTRO_TELEMETRY_DISABLED: '1', NODE_OPTIONS: `${process.env.NODE_OPTIONS || ''} --import=${guard}`.trim() },
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
