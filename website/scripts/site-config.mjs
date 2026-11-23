import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const website = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const repository = path.dirname(website);
const trimmed = (process.env.DOCS_BASE ?? '/RATools-for-PDF').trim().replace(/^\/+|\/+$/g, '');
if (trimmed && !trimmed.split('/').every((part) => /^[A-Za-z0-9_-]+$/.test(part))) {
  throw new Error('DOCS_BASE must be a path such as /RATools-for-PDF or /pdf.');
}
export const base = trimmed ? `/${trimmed}` : '';
export const site = process.env.DOCS_SITE || 'https://PharmaRA.github.io';
const origin = new URL(site);
if (!['http:', 'https:'].includes(origin.protocol) || origin.username || origin.password ||
    origin.pathname !== '/' || origin.search || origin.hash) {
  throw new Error('DOCS_SITE must be an HTTP(S) origin without credentials or a path.');
}
export const repoUrl = 'https://github.com/PharmaRA/RATools-for-PDF';
