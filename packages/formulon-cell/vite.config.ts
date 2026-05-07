import { resolve } from 'node:path';
import dts from 'unplugin-dts/vite';
import { defineConfig } from 'vite';

// Library build for `@libraz/formulon-cell`.
//
// Output uses `preserveModules` so dist/ mirrors src/. That matters because
// `engine/loader.ts` references the WASM via `new URL('../../vendor/...',
// import.meta.url)` — the same relative path has to resolve in both the
// dev workspace layout and the published-package layout. Mirroring src
// gives us that for free.
export default defineConfig({
  build: {
    target: 'es2022',
    sourcemap: true,
    assetsInlineLimit: 0,
    minify: false, // tree-shaken consumers do their own minification
    lib: {
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        'extensions/index': resolve(__dirname, 'src/extensions/index.ts'),
        'i18n/ja': resolve(__dirname, 'src/i18n/ja.ts'),
        'i18n/en': resolve(__dirname, 'src/i18n/en.ts'),
      },
      formats: ['es'],
    },
    rollupOptions: {
      // The browser Emscripten bundle uses TLA + new URL(...) for both the
      // wasm and worker scripts. Vite's lib worker plugin tries to bundle
      // those as iife workers, which breaks on TLA. Externalizing keeps the
      // import alive — at consume time the user's bundler resolves
      // `../../vendor/formulon/formulon.web.js` relative to our shipped
      // `dist/engine/loader.js` and handles the Emscripten module on its own.
      external: [
        'zustand',
        'zustand/vanilla',
        /\/vendor\/formulon\/formulon(?:\.web)?\.js$/,
      ],
      output: {
        preserveModules: true,
        preserveModulesRoot: 'src',
        entryFileNames: '[name].js',
      },
    },
  },
  plugins: [
    dts({
      // Use a build-only tsconfig that pins `rootDir: ./src` and excludes
      // tests/. TS 6 made `rootDir` default to the tsconfig directory, so
      // the regular tsconfig (which includes tests/) would otherwise emit
      // .d.ts under dist/src/, breaking the package's exports map.
      tsconfigPath: './tsconfig.build.json',
      // No rollupTypes — we want per-file .d.ts so subpath imports get
      // their own type declarations.
    }),
  ],
});
