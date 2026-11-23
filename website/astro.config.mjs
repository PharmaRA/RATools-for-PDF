import { defineConfig } from 'astro/config';
import starlight from '@astrojs/starlight';
import { readFileSync } from 'node:fs';
import { base, site } from './scripts/site-config.mjs';

const sidebar = JSON.parse(readFileSync(new URL('./src/data/sidebar.json', import.meta.url), 'utf8'));

export default defineConfig({
  site,
  base: base || '/',
  trailingSlash: 'always',
  output: 'static',
  // The guarded runner recycles the previous output before building.
  vite: { build: { emptyOutDir: false } },
  integrations: [
    starlight({
      title: 'RATools for PDF',
      description: 'PDF 批量处理中文使用手册：预检、页面与字体、书签与链接、压缩、安全中心和结果复核。',
      favicon: '/favicon.svg',
      defaultLocale: 'root',
      locales: { root: { label: '简体中文', lang: 'zh-CN' } },
      social: [{ icon: 'github', label: 'GitHub', href: 'https://github.com/PharmaRA/RATools-for-PDF' }],
      customCss: ['./src/styles/theme.css'],
      components: { Header: './src/components/SiteHeader.astro' },
      sidebar,
      pagefind: true,
      tableOfContents: { minHeadingLevel: 2, maxHeadingLevel: 3 },
      credits: false,
      head: [
        { tag: 'meta', attrs: { name: 'theme-color', content: '#1769d2' } },
        { tag: 'meta', attrs: { property: 'og:locale', content: 'zh_CN' } },
      ],
    }),
  ],
});
