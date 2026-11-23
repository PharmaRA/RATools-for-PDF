import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { base, website, repository, repoUrl } from './site-config.mjs';

const source = path.join(repository, 'docs/manual');
const target = path.join(website, 'src/content/docs/manual');
const assetRoot = path.join(website, 'public/manual-assets');

function inside(parent, child) {
  const relative = path.relative(parent, child);
  return relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function assertPlain(file) {
  let current = file;
  while (inside(repository, current)) {
    if (fs.existsSync(current)) {
      const stat = fs.lstatSync(current);
      if (stat.isSymbolicLink() || (stat.isFile() && stat.nlink > 1)) {
        throw new Error(`Linked path is not supported: ${path.relative(repository, current)}`);
      }
    }
    if (current === repository) break;
    current = path.dirname(current);
  }
}

function slugs(nodes) {
  return nodes.flatMap((node) => node.slug ? [node.slug] : slugs(node.items ?? []));
}

function writeChanged(file, content) {
  assertPlain(file);
  const bytes = Buffer.isBuffer(content) ? content : Buffer.from(content);
  if (fs.existsSync(file) && fs.readFileSync(file).equals(bytes)) return;
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, bytes);
}

function filesBelow(directory) {
  if (!fs.existsSync(directory)) return [];
  assertPlain(directory);
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const file = path.join(directory, entry.name);
    assertPlain(file);
    return entry.isDirectory() ? filesBelow(file) : [file];
  });
}

export function syncDocs({ quiet = false } = {}) {
  assertPlain(source);
  const pages = fs.readdirSync(source).filter((name) => name.endsWith('.md')).sort();
  const expected = new Set(pages.map((name) => name === 'index.md' ? 'manual' : `manual/${name.slice(0, -3)}`));
  const sidebar = JSON.parse(fs.readFileSync(path.join(website, 'src/data/sidebar.json'), 'utf8'));
  const configured = slugs(sidebar);
  if (configured.length !== expected.size || new Set(configured).size !== expected.size ||
      configured.some((slug) => !expected.has(slug))) {
    throw new Error('Sidebar must list every manual page exactly once.');
  }
  const planned = new Map();
  const assets = new Map();
  let rewrittenLinks = 0;
  for (const name of pages) {
    const file = path.join(source, name);
    assertPlain(file);
    const text = fs.readFileSync(file, 'utf8');
    const title = text.match(/^#\s+(.+)$/m)?.[1];
    if (!title) throw new Error(`Missing title: ${name}`);
    let fenced = false;
    const body = text.replace(/^#\s+.+\r?\n/, '').trimStart().split(/\r?\n/).map((line) => {
      if (/^\s*```/.test(line)) { fenced = !fenced; return line; }
      if (fenced) return line;
      return line.replace(/(!?\[[^\]]*\])\(([^)\s]+)\)/g, (whole, label, href) => {
        if (/^(?:https?:|mailto:|#)/i.test(href)) return whole;
        const parsed = new URL(href, pathToFileURL(file));
        if (parsed.protocol !== 'file:') throw new Error(`Unsupported local link in ${name}`);
        const resolved = path.resolve(path.dirname(file), decodeURIComponent(href.split(/[?#]/)[0]));
        if (!inside(repository, resolved) || !fs.existsSync(resolved)) {
          throw new Error(`Missing or out-of-repository target in ${name}: ${href}`);
        }
        assertPlain(resolved);
        let rewritten;
        if (path.dirname(resolved) === source && resolved.endsWith('.md')) {
          const stem = path.basename(resolved, '.md');
          rewritten = `${base}/manual/${stem === 'index' ? '' : `${stem}/`}`;
        } else if (label.startsWith('!')) {
          if (!inside(path.join(repository, 'docs/assets'), resolved)) {
            throw new Error(`Unexpected manual image: ${href}`);
          }
          const relative = path.relative(path.join(repository, 'docs/assets'), resolved).replaceAll('\\', '/');
          assets.set(path.join(assetRoot, relative), fs.readFileSync(resolved));
          rewritten = `${base}/manual-assets/${relative}`;
        } else {
          const relative = path.relative(repository, resolved).replaceAll('\\', '/');
          rewritten = `${repoUrl}/blob/main/${relative.split('/').map(encodeURIComponent).join('/')}`;
        }
        rewrittenLinks++;
        return `${label}(${rewritten}${parsed.search}${parsed.hash})`;
      });
    }).join('\n').replace(/(<a id="[^"]+"><\/a>)\n(##)/g, '$1\n\n$2');
    const paragraph = body.split('\n\n').find((part) => part.trim() && !/^(?:\[返回|#|<|\|)/.test(part.trim())) ?? title;
    const description = paragraph.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1').replace(/[*`]/g, '').replace(/\s+/g, ' ').slice(0, 155);
    const front = [
      '---', `title: ${JSON.stringify(title)}`, `description: ${JSON.stringify(description)}`,
      `editUrl: ${JSON.stringify(`${repoUrl}/edit/main/docs/manual/${name}`)}`, '---', '',
    ].join('\n');
    planned.set(path.join(target, name), `${front}\n${body}`);
  }
  // Refuse stale mirrors rather than silently deleting source or generated files.
  for (const [directory, files] of [[target, planned], [assetRoot, assets]]) {
    for (const existing of filesBelow(directory)) {
      if (!files.has(existing)) throw new Error(`Stale generated file; move it to the Recycle Bin: ${path.relative(website, existing)}`);
    }
  }
  for (const [file, content] of [...planned, ...assets]) writeChanged(file, content);
  if (!quiet) console.log(`Synced ${pages.length} pages, ${assets.size} assets, ${rewrittenLinks} local links; base=${base || '/'}`);
  return { pages: pages.length, assets: assets.size };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  syncDocs();
}
