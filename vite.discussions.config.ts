import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';
import { readFileSync } from 'node:fs';

export default defineConfig({
  plugins: [react(), {
    name: 'discussion-only-package',
    generateBundle(_options, bundle) {
      const forbidden = /\/src\/(?:background|popup|access-review|no-review|import-review)\/|\/src\/oneoffice\/(?:noWrite|noUpdate|projectImport|projectCodeNo|testBaseline|testPreservation|testTaskDetails|accessHistory|accessPersonnel)|\/src\/shared\/(?:storage|reportGenerator|templateGenerator)\./;
      for (const item of Object.values(bundle)) if (item.type === 'chunk') {
        for (const module of Object.keys(item.modules)) if (forbidden.test(module)) throw new Error(`Non-discussion module in package: ${module}`);
      }
      const manifest = { manifest_version: 3, name: 'T-eraClip Thảo luận BETA', version: '0.1.0',
        description: 'Bản thử nghiệm thu thập và xuất thảo luận 1Office của một công việc hoặc toàn dự án. THIS EXTENSION IS FOR BETA TESTING.',
        permissions: ['storage', 'activeTab', 'scripting'], action: { default_popup: 'discussion-popup.html', default_title: 'T-eraClip Thảo luận', default_icon: { '16': 'icons/icon-16.png', '32': 'icons/icon-32.png' } },
        icons: { '16': 'icons/icon-16.png', '32': 'icons/icon-32.png', '48': 'icons/icon-48.png', '128': 'icons/icon-128.png' } };
      this.emitFile({ type: 'asset', fileName: 'manifest.json', source: JSON.stringify(manifest, null, 2) });
      for (const size of [16, 32, 48, 128]) this.emitFile({ type: 'asset', fileName: `icons/icon-${size}.png`, source: readFileSync(resolve(__dirname, `public/icons/icon-${size}.png`)) });
      this.emitFile({ type: 'asset', fileName: 'tadt-logo.png', source: readFileSync(resolve(__dirname, 'public/tadt-logo.png')) });
    },
  }],
  publicDir: false,
  build: { outDir: 'dist-discussions', rollupOptions: { input: { popup: resolve(__dirname, 'discussion-popup.html'), review: resolve(__dirname, 'review.html'), projectReview: resolve(__dirname, 'project-review.html') },
    output: { entryFileNames: '[name].js', chunkFileNames: 'assets/[name]-[hash].js', assetFileNames: 'assets/[name]-[hash][extname]' } } },
});
