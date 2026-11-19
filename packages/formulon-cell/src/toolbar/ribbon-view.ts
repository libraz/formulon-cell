import { spreadsheetFailure } from '../api/errors.js';
import type { Disposer, Emitter } from '../api/events.js';
import { canExecuteBuiltInRequest } from '../commands/built-in-command-policy.js';
import type { ToolbarInstance } from '../mount/toolbar-types.js';
import type { SpreadsheetInstance as LegacyInstance } from '../mount/types.js';
import type { RibbonDisplayMode } from './ribbon/render-ribbon.js';
import type { RibbonTab } from './ribbon-model.js';

/** What `update` on a ribbon can change. */
export interface ToolbarPatch {
  readonly activeTab?: RibbonTab;
  readonly displayMode?: RibbonDisplayMode;
  readonly formulaBarVisible?: boolean;
  readonly backstageOpen?: boolean;
}

/** The state of a ribbon, as `change` reports it. */
export interface ToolbarState {
  readonly activeTab: RibbonTab;
  readonly displayMode: RibbonDisplayMode;
  readonly formulaBarVisible: boolean;
  readonly backstageOpen: boolean;
}

/** Events of a ribbon. */
export interface RibbonEvents {
  change: ToolbarState;
}

/** The ribbon of an instance, `SpreadsheetInstance.toolbar`: its state and commands. It comes
 *  and goes with `update({ toolbar })`; a ribbon placed on its own is a `Toolbar`. */
export interface InstanceRibbon extends ToolbarState {
  /** Applies the given state. */
  update(patch: ToolbarPatch): void;
  /**
   * Runs a ribbon command on the current selection; throws `policy-denied` when the policy refuses it.
   *
   * Accepts the id of a ribbon button and the id of an item in a button's menu. Only button ids are
   * `CommandId`s that `SpreadsheetInstance.commands` runs; menu items run here.
   */
  execute(commandId: string): void;
  /** Moves focus to the active tab; false when the ribbon has none to focus. */
  focusActiveTab(): boolean;
  /** Subscribes to state changes. */
  on(name: 'change', listener: (state: ToolbarState) => void): Disposer;
}

/** The state a ribbon handle shows now. */
export const ribbonStateOf = (handle: ToolbarInstance): ToolbarState => ({
  activeTab: handle.getActiveTab(),
  displayMode: handle.getDisplayMode(),
  formulaBarVisible: handle.getFormulaBarVisible(),
  backstageOpen: handle.getBackstageOpen(),
});

/** An {@link InstanceRibbon} over a ribbon handle; `assertLive` guards every member. */
export class RibbonView implements InstanceRibbon {
  readonly #handle: ToolbarInstance;
  readonly #legacy: LegacyInstance;
  readonly #events: Emitter<RibbonEvents>;
  readonly #assertLive: () => void;

  constructor(
    handle: ToolbarInstance,
    legacy: LegacyInstance,
    events: Emitter<RibbonEvents>,
    assertLive: () => void,
  ) {
    this.#handle = handle;
    this.#legacy = legacy;
    this.#events = events;
    this.#assertLive = assertLive;
  }

  get activeTab(): RibbonTab {
    this.#assertLive();
    return this.#handle.getActiveTab();
  }

  get displayMode(): RibbonDisplayMode {
    this.#assertLive();
    return this.#handle.getDisplayMode();
  }

  get formulaBarVisible(): boolean {
    this.#assertLive();
    return this.#handle.getFormulaBarVisible();
  }

  get backstageOpen(): boolean {
    this.#assertLive();
    return this.#handle.getBackstageOpen();
  }

  update(patch: ToolbarPatch): void {
    this.#assertLive();
    const h = this.#handle;
    if (patch.displayMode !== undefined) h.setDisplayMode(patch.displayMode);
    if (patch.formulaBarVisible !== undefined) h.setFormulaBarVisible(patch.formulaBarVisible);
    if (patch.activeTab !== undefined) h.setActiveTab(patch.activeTab);
    if (patch.backstageOpen !== undefined) h.setBackstageOpen(patch.backstageOpen);
  }

  execute(commandId: string): void {
    this.#assertLive();
    if (!canExecuteBuiltInRequest(this.#legacy.store, commandId, 'ribbon').allowed) {
      throw spreadsheetFailure('policy-denied', `command ${commandId} may not run`, {
        context: { command: commandId },
      });
    }
    this.#handle.applyCommand(commandId);
  }

  focusActiveTab(): boolean {
    this.#assertLive();
    return this.#handle.focusActiveTab();
  }

  on(name: 'change', listener: (state: ToolbarState) => void): Disposer {
    this.#assertLive();
    return this.#events.on(name, listener);
  }
}
