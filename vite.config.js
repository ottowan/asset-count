import { fileURLToPath, URL } from 'node:url';
import { defineConfig } from 'vite';

export default defineConfig({
  resolve: {
    alias: {
      '@zxing/browser': fileURLToPath(new URL('./node_modules/@zxing/browser/esm/index.js', import.meta.url)),
    },
  },
});
