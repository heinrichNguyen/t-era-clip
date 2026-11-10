import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'path';

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        popup: resolve(__dirname, 'popup.html'),
        options: resolve(__dirname, 'options.html'),
        review: resolve(__dirname, 'review.html'),
        projectReview: resolve(__dirname, 'project-review.html'),
        accessReview: resolve(__dirname, 'access-review.html'),
        noReview: resolve(__dirname, 'no-review.html'),
        importReview: resolve(__dirname, 'import-review.html'),
        background: resolve(__dirname, 'src/background/index.ts')
      },
      output: {
        entryFileNames: '[name].js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash][extname]'
      }
    }
  }
});
