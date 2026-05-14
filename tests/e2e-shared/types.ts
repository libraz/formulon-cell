/**
 * App-agnostic descriptor for the three demo apps. Used by the shared
 * Playwright config + scenarios so the same spec can run in playground,
 * react-demo, and vue-demo.
 */
export interface DemoApp {
  /** Human-readable id used in test titles + artifact paths. */
  id: 'playground' | 'react-demo' | 'vue-demo';
  /** TCP port the dev server binds to. Mirrors apps/<id>/vite.config.ts. */
  port: number;
  /** Yarn workspace name used by `yarn workspace <name> dev`. */
  workspace: string;
}

export const DEMO_APPS: Readonly<Record<DemoApp['id'], DemoApp>> = Object.freeze({
  playground: { id: 'playground', port: 5173, workspace: '@formulon-cell/playground' },
  'react-demo': { id: 'react-demo', port: 5174, workspace: '@formulon-cell/react-demo' },
  'vue-demo': { id: 'vue-demo', port: 5175, workspace: '@formulon-cell/vue-demo' },
});
