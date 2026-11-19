import { mountToolbar } from './toolbar.js';
import type { ToolbarInstance } from './toolbar-types.js';
import type { SpreadsheetInstance } from './types.js';

export interface RibbonHostDeps {
  host: HTMLElement;
  getInstance: () => SpreadsheetInstance;
  getLocale: () => string;
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
        dynamicDropdowns: true,
        ...toolbarOpts,
      });
    },
    getToolbar: () => toolbarHandle,
    dispose,
  };
}
