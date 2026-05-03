import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import dts from 'vite-plugin-dts';

export default defineConfig({
  build: {
    target: 'es2022',
    sourcemap: true,
    assetsInlineLimit: 0,
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      formats: ['es'],
      fileName: () => 'index.js',
    },
    rollupOptions: {
      external: ['zustand', 'zustand/vanilla'],
    },
  },
  assetsInclude: ['**/*.wasm'],
  plugins: [
    dts({
      rollupTypes: true,
      tsconfigPath: './tsconfig.json',
      include: ['src/**/*'],
    }),
  ],
});
