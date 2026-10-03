import { mountToolbar, type ToolbarInstance, type ToolbarInstanceRef } from './toolbar.js';
import type { SpreadsheetInstance } from './types.js';

export interface RibbonHostDeps {
  host: HTMLElement;
  getInstance: () => ToolbarInstanceRef;
  getLocale: () => string;
  /** True while an interaction policy is active (dynamic dropdowns stay off). */
  isRestricted: () => boolean;
}

export interface RibbonHost {
  /** Rebuild the ribbon for `next`; a falsy value only tears it down. */
  setToolbar: (next: Parameters<SpreadsheetInstance['setToolbar']>[0]) => void;
  getToolbar: () => ToolbarInstance | null;
  dispose: () => void;
}

export function createRibbonHost(deps: RibbonHostDeps): RibbonHost {
  const { host } = deps;
  let toolbarHandle: ToolbarInstance | null = null;
  let ribbonHost: HTMLElement | null = null;
  const dispose = (): void => {
    toolbarHandle?.dispose();
    toolbarHandle = null;
    ribbonHost?.remove();
    ribbonHost = null;
  };
  return {
    setToolbar(next) {
      dispose();
      if (!next) return;
      ribbonHost = host.ownerDocument.createElement('div');
      ribbonHost.className = 'fc-host__ribbon';
      ribbonHost.style.display = 'contents';
      host.insertBefore(ribbonHost, host.firstChild);
      const toolbarOpts = next === true ? {} : next;
      toolbarHandle = mountToolbar(ribbonHost, deps.getInstance(), {
        lang: deps.getLocale() === 'en' ? 'en' : 'ja',
        ...(deps.isRestricted() ? {} : { dynamicDropdowns: true as const }),
        ...toolbarOpts,
      });
    },
    getToolbar: () => toolbarHandle,
    dispose,
  };
}
