import { defineConfig } from 'vite';

export default defineConfig({
  // Relative base ensures asset paths work on GitHub Pages (<user>.github.io/<repo>/)
  base: './',
  build: {
    outDir: 'dist',
    rollupOptions: {
      output: {
        manualChunks: {
          vendor: ['tabulator-tables', 'lucide'],
          xlsx: ['xlsx']
        }
      }
    }
  }
});

