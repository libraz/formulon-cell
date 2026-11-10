/**
 * Specs a secondary demo app runs by default.
 *
 * The demos share every scenario, so running all of them through both adapters
 * repeats the same core behaviour twice. Locally, the primary app runs the full
 * set and a secondary app runs only the specs below: the ones that pass through
 * adapter-specific code (component mount, props, events, the toolbar component,
 * locale binding) or through the framework-specific demo shell. CI and
 * `FC_E2E_FULL=1` run every spec on every app.
 */
export const PRIMARY_DEMO_APP = 'react-demo';

export const ADAPTER_SPECS: readonly { readonly spec: string; readonly reason: string }[] = [
  { spec: 'wrapper.spec.ts', reason: 'adapter component mount, props and events' },
  { spec: 'smoke.spec.ts', reason: 'boots through the adapter' },
  { spec: 'i18n.spec.ts', reason: 'locale and strings props through the adapter' },
  { spec: 'error.spec.ts', reason: 'mount failure reporting through the adapter' },
  { spec: 'restricted-embedding.spec.ts', reason: 'policy and viewport props through the adapter' },
  { spec: 'excel-chrome.spec.ts', reason: 'toolbar adapter component' },
  { spec: 'mac-chrome.spec.ts', reason: 'toolbar adapter component on the Mac profile' },
  { spec: 'a11y-ribbon.spec.ts', reason: 'toolbar adapter component accessibility' },
  { spec: 'ribbon-tab-focus.spec.ts', reason: 'toolbar adapter component focus' },
  { spec: 'demo-modal-focus.spec.ts', reason: 'framework demo shell modals' },
  { spec: 'host-modal-collision.spec.ts', reason: 'framework demo shell modals over the grid' },
  { spec: 'user-journeys.spec.ts', reason: 'end-to-end journeys through the demo shell' },
];

/** True when every spec should run on every app. */
export const runFullE2E = (): boolean => !!process.env.CI || process.env.FC_E2E_FULL === '1';
