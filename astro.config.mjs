// @ts-check
import { defineConfig } from 'astro/config';
import tailwindcss from '@tailwindcss/vite';
import sitemap from '@astrojs/sitemap';
import vb from '@shikijs/langs/vb';

// https://astro.build/config
export default defineConfig({
  site: 'https://decodingvba.com',
  integrations: [sitemap()],
  markdown: {
    shikiConfig: {
      langs: [vb],
    },
  },
  vite: {
    plugins: [tailwindcss()],
  },
});
