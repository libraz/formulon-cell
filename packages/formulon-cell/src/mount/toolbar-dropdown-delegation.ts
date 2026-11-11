// Document-level delegation for the ribbon's dynamic dropdowns. The listeners
// sit on `document` so clicks anywhere inside an open `.fc-tb__menu` reach the
// dispatcher; `detach` removes every one of them on toolbar dispose.

import { interactionControllerFor } from '../commands/interaction-controller.js';
import { delegatePointerArrive } from '../toolbar/menu-a11y.js';
import type { RibbonHooks } from '../toolbar/ribbon/apply-ribbon-command.js';
import {
  createDynamicDropdowns,
  type DynamicDropdownsApi,
  type DynamicDropdownsCtx,
} from '../toolbar/ribbon/dynamic-dropdowns.js';
import { MAC_FORMULAS_MORE_MENU_ID } from '../toolbar/ribbon/mac/model.js';
import { createDefaultDynamicDropdownsCtx } from './dynamic-dropdowns-defaults.js';
import type { SpreadsheetInstance } from './types.js';

type DynamicDropdownsOption =
  | true
  | Partial<DynamicDropdownsCtx>
  | (() => Partial<DynamicDropdownsCtx>);

export interface DynamicDropdownDelegationDeps {
  defaultsInstance: SpreadsheetInstance | null;
  getInstance: () => SpreadsheetInstance | null;
  focusSheet: () => void;
  projectFormatToolbar: () => void;
  refreshCells: () => void;
  closeBorderMenu: (restoreFocus?: boolean) => void;
  /** Host hooks as passed to the toolbar; read at call time. */
  getHostHooks?: () => RibbonHooks | undefined;
  /** Called after a click or Enter/Space dispatch left the menu closed. */
  onMenuClosedByCommand: () => void;
}

export interface DynamicDropdownDelegation {
  api: DynamicDropdownsApi;
  detach: () => void;
}

export function attachDynamicDropdownDelegation(
  option: DynamicDropdownsOption,
  deps: DynamicDropdownDelegationDeps,
): DynamicDropdownDelegation {
  const { defaultsInstance, getInstance, onMenuClosedByCommand } = deps;
  const hostOverrides: Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>) =
    option === true ? {} : option;
  let coreDefaults: DynamicDropdownsCtx | null = null;
  const getCoreDefaults = (): DynamicDropdownsCtx => {
    coreDefaults ??= createDefaultDynamicDropdownsCtx(
      (defaultsInstance ?? ({} as SpreadsheetInstance)) as unknown as Parameters<
        typeof createDefaultDynamicDropdownsCtx
      >[0],
    );
    return coreDefaults;
  };
  // The host's automation hooks take the custom-script and add-in manager items;
  // every other action keeps the core default.
  const hookRouting = (): Partial<DynamicDropdownsCtx> => {
    const automation = deps.getHostHooks?.()?.automation;
    const routed: Partial<DynamicDropdownsCtx> = {};
    if (automation?.runScript) {
      routed.applyScriptAction = (action) =>
        action === 'custom'
          ? void deps.getHostHooks?.()?.automation?.runScript?.()
          : getCoreDefaults().applyScriptAction(action);
    }
    if (automation?.addInManager) {
      routed.applyAddInAction = (action) =>
        action === 'manage'
          ? deps.getHostHooks?.()?.automation?.addInManager?.()
          : getCoreDefaults().applyAddInAction(action);
    }
    return routed;
  };
  const withToolbarDropdownOverrides = (
    overrides: Partial<DynamicDropdownsCtx>,
  ): Partial<DynamicDropdownsCtx> => ({
    closeBorderMenu: deps.closeBorderMenu,
    ...hookRouting(),
    ...overrides,
  });
  const overridesOpt: Partial<DynamicDropdownsCtx> | (() => Partial<DynamicDropdownsCtx>) =
    typeof hostOverrides === 'function'
      ? () => withToolbarDropdownOverrides(hostOverrides())
      : withToolbarDropdownOverrides(hostOverrides);
  // `createDefaultDynamicDropdownsCtx` uses the `@libraz/formulon-cell`
  // self-import for `SpreadsheetInstance` (matching `dynamic-dropdowns.ts`)
  // so its parameter type resolves to dist. This file imports the
  // src-side declaration via `./types.js`, so the two structurally
  // identical declarations need one bridge cast.
  //
  // When `defaultsInstance` is null (deferred-mount hosts like the
  // playground), the built-in base handlers stay unreachable as long as
  // the host overrides every handler it dispatches. The override getter
  // (recommended for deferred hosts) captures the live instance via its
  // own closure so it can hand back the real `inst` once mounted.
  const dropdownsCtx = createDefaultDynamicDropdownsCtx(
    (defaultsInstance ?? ({} as SpreadsheetInstance)) as unknown as Parameters<
      typeof createDefaultDynamicDropdownsCtx
    >[0],
    {
      focusSheet: deps.focusSheet,
      projectFormatToolbar: deps.projectFormatToolbar,
      refreshCells: deps.refreshCells,
      overrides: overridesOpt,
    },
  );
  const api = createDynamicDropdowns(dropdownsCtx);

  const onClick = (event: MouseEvent): void => {
    const current = getInstance();
    const target = event.target instanceof Element ? event.target : null;
    const menu = target?.closest<HTMLElement>('.fc-tb__menu') ?? null;
    if (
      current &&
      interactionControllerFor(current.store)?.policy !== undefined &&
      !target?.closest<HTMLElement>(
        `#${MAC_FORMULAS_MORE_MENU_ID} [data-function-category-submenu]`,
      )
    )
      return;
    if (api.dynamicRibbonDropdownClick(event) && menu?.hidden) onMenuClosedByCommand();
  };
  const onPointerDown = (event: MouseEvent): void => {
    api.dynamicRibbonDropdownPointerDown(event);
  };
  const onFocusIn = (event: FocusEvent): void => {
    api.dynamicRibbonDropdownFocusIn(event);
  };
  const onHover = (event: MouseEvent): void => {
    api.dynamicRibbonDropdownHover(event);
  };
  const onKeyDown = (event: KeyboardEvent): void => {
    const current = getInstance();
    const target = event.target instanceof Element ? event.target : null;
    const menu = target?.closest<HTMLElement>('.fc-tb__menu') ?? null;
    if (
      current &&
      interactionControllerFor(current.store)?.policy !== undefined &&
      !target?.closest<HTMLElement>(`#${MAC_FORMULAS_MORE_MENU_ID}`)
    )
      return;
    if (
      api.dynamicRibbonDropdownKeydown(event) &&
      menu?.hidden &&
      (event.key === 'Enter' || event.key === ' ')
    )
      onMenuClosedByCommand();
  };
  document.addEventListener('click', onClick);
  document.addEventListener('mousedown', onPointerDown, true);
  document.addEventListener('focusin', onFocusIn);
  const detachHover = delegatePointerArrive(document, onHover);
  document.addEventListener('keydown', onKeyDown);

  return {
    api,
    detach: () => {
      document.removeEventListener('click', onClick);
      document.removeEventListener('mousedown', onPointerDown, true);
      document.removeEventListener('focusin', onFocusIn);
      detachHover();
      document.removeEventListener('keydown', onKeyDown);
    },
  };
}
