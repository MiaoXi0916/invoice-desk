import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';

export default defineConfig({
  root: __dirname,
  plugins: [react()],
  base: './',
  build: { outDir: path.resolve(__dirname, '..', '发布产物', 'dist'), emptyOutDir: true }
});
