/**
 * App-agnostic descriptor for the framework demo apps. Used by the shared
 * Playwright config + scenarios so the same spec can run in react-demo and
 * vue-demo.
 */
export interface DemoApp {
  /** Human-readable id used in test titles + artifact paths. */
  id: 'react-demo' | 'vue-demo';
  /** TCP port the dev server binds to. Mirrors apps/<id>/vite.config.ts. */
  port: number;
  /** Yarn workspace name used by `yarn workspace <name> dev`. */
  workspace: string;
}

export const DEMO_APPS: Readonly<Record<DemoApp['id'], DemoApp>> = Object.freeze({
  'react-demo': { id: 'react-demo', port: 5174, workspace: '@formulon-cell/react-demo' },
  'vue-demo': { id: 'vue-demo', port: 5175, workspace: '@formulon-cell/vue-demo' },
});
