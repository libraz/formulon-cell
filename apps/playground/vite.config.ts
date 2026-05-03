import { defineConfig } from 'vite';

export default defineConfig({
  server: {
    port: 5173,
    // formulon ships a pthread-enabled WASM that uses SharedArrayBuffer.
    // Browsers require crossOriginIsolated context (COOP+COEP) to allow it.
    // `credentialless` keeps Google Fonts / unpkg-style CDNs working without
    // forcing every third-party asset to expose Cross-Origin-Resource-Policy.
    headers: {
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
    },
    fs: {
      // Allow serving the workspace's vendored WASM.
      allow: ['..', '../..'],
    },
  },
  optimizeDeps: {
    // Don't try to pre-bundle the formulon emscripten module — it doesn't
    // play well with esbuild's CJS interop.
    exclude: ['@formulon-cell/core'],
  },
});
