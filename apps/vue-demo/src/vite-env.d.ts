/// <reference types="vite/client" />

declare module '*.vue' {
  import type { DefineComponent } from 'vue';
  // biome-ignore lint/complexity/noBannedTypes: Vue's DefineComponent expects {}
  const component: DefineComponent<{}, {}, unknown>;
  export default component;
}
