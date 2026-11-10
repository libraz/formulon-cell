import { spreadsheetFailure } from '../api/errors.js';
import { type Disposer, Emitter } from '../api/events.js';
import { type Disposable, Lifecycle } from '../api/lifecycle.js';
import { legacyInstanceOf, type SpreadsheetInstance } from '../api/spreadsheet/instance.js';
import { canExecuteBuiltInRequest } from '../commands/built-in-command-policy.js';
import {
  type MountToolbarOptions as LegacyToolbarOptions,
  mountToolbar as mountLegacyToolbar,
  type ToolbarInstance,
} from '../mount/toolbar.js';
import type { SpreadsheetInstance as LegacyInstance } from '../mount/types.js';
import type { RibbonHooks } from './ribbon/apply-ribbon-command.js';
import type { RibbonDisplayMode } from './ribbon/render-ribbon.js';
import type { RibbonTab } from './ribbon-model.js';

/** Options of a ribbon, for `ribbon(options)` and `mountToolbar(host, instance, options)`. */
export interface RibbonOptions {
  /** Host actions for the commands the ribbon cannot run on its own. */
  readonly hooks?: RibbonHooks;
  /** The tab surface; defaults to the platform's standard tabs. */
  readonly ribbonTabs?: readonly RibbonTab[];
  readonly activeTab?: RibbonTab;
  readonly displayMode?: RibbonDisplayMode;
  readonly formulaBarVisible?: boolean;
}

/** Options of a ribbon placed on its own. */
export interface MountToolbarOptions extends RibbonOptions {
  /** Aborting rejects with `signal.reason` before `host` is touched. */
  readonly signal?: AbortSignal;
}

/** What `Toolbar.update` can change. */
export interface ToolbarPatch {
  readonly activeTab?: RibbonTab;
  readonly displayMode?: RibbonDisplayMode;
  readonly formulaBarVisible?: boolean;
  readonly backstageOpen?: boolean;
}

/** The state of a toolbar, as `change` reports it. */
export interface ToolbarState {
  readonly activeTab: RibbonTab;
  readonly displayMode: RibbonDisplayMode;
  readonly formulaBarVisible: boolean;
  readonly backstageOpen: boolean;
}

interface ToolbarEvents {
  change: ToolbarState;
}

/** A ribbon mounted on its own with `mountToolbar`. */
export interface Toolbar extends Disposable, ToolbarState {
  /** Applies the given state; the ribbon of an instance is configured with `update({ toolbar })`. */
  update(patch: ToolbarPatch): void;
  /** Runs a ribbon command on the current selection; throws `policy-denied` when the policy refuses it. */
  execute(commandId: string): void;
  /** Moves focus to the active tab; false when the ribbon has none to focus. */
  focusActiveTab(): boolean;
  /** Subscribes to state changes. */
  on(name: 'change', listener: (state: ToolbarState) => void): Disposer;
}

class ToolbarImpl implements Toolbar {
  readonly #life = new Lifecycle('Toolbar');
  readonly #events: Emitter<ToolbarEvents>;
  readonly #handle: ToolbarInstance;
  readonly #legacy: LegacyInstance;

  constructor(
    handle: ToolbarInstance,
    legacy: LegacyInstance,
    events: Emitter<ToolbarEvents>,
    release: () => void,
  ) {
    this.#events = events;
    this.#handle = handle;
    this.#legacy = legacy;
    this.#life.onDispose(() => {
      this.#events.clear();
      release();
      handle.dispose();
    });
  }

  get disposed(): boolean {
    return this.#life.disposed;
  }

  get activeTab(): RibbonTab {
    this.#life.assertLive();
    return this.#handle.getActiveTab();
  }

  get displayMode(): RibbonDisplayMode {
    this.#life.assertLive();
    return this.#handle.getDisplayMode();
  }

  get formulaBarVisible(): boolean {
    this.#life.assertLive();
    return this.#handle.getFormulaBarVisible();
  }

  get backstageOpen(): boolean {
    this.#life.assertLive();
    return this.#handle.getBackstageOpen();
  }

  update(patch: ToolbarPatch): void {
    this.#life.assertLive();
    const h = this.#handle;
    if (patch.displayMode !== undefined) h.setDisplayMode(patch.displayMode);
    if (patch.formulaBarVisible !== undefined) h.setFormulaBarVisible(patch.formulaBarVisible);
    if (patch.activeTab !== undefined) h.setActiveTab(patch.activeTab);
    if (patch.backstageOpen !== undefined) h.setBackstageOpen(patch.backstageOpen);
  }

  execute(commandId: string): void {
    this.#life.assertLive();
    if (!canExecuteBuiltInRequest(this.#legacy.store, commandId, 'ribbon').allowed) {
      throw spreadsheetFailure('policy-denied', `command ${commandId} may not run`, {
        context: { command: commandId },
      });
    }
    this.#handle.applyCommand(commandId);
  }

  focusActiveTab(): boolean {
    this.#life.assertLive();
    return this.#handle.focusActiveTab();
  }

  on(name: 'change', listener: (state: ToolbarState) => void): Disposer {
    this.#life.assertLive();
    return this.#events.on(name, listener);
  }

  dispose(): void {
    this.#life.dispose();
  }

  [Symbol.dispose](): void {
    this.dispose();
  }
}

/**
 * Mounts a ribbon into `host`, apart from the grid, driving `instance`. Rejects with
 * `signal.reason` when `signal` is aborted before or while it mounts.
 */
export async function mountToolbar(
  host: HTMLElement,
  instance: SpreadsheetInstance,
  options: MountToolbarOptions = {},
): Promise<Toolbar> {
  const { signal } = options;
  signal?.throwIfAborted();
  const legacy = legacyInstanceOf(instance);
  const themeOf = (): NonNullable<LegacyToolbarOptions['theme']> =>
    legacy.store.getState().ui.theme as NonNullable<LegacyToolbarOptions['theme']>;
  const previousTheme = host.dataset.fcTheme;
  host.dataset.fcTheme = themeOf();
  const events = new Emitter<ToolbarEvents>();
  let handle: ToolbarInstance | null = null;
  const notify = (): void => {
    if (!handle) return;
    events.emit('change', {
      activeTab: handle.getActiveTab(),
      displayMode: handle.getDisplayMode(),
      formulaBarVisible: handle.getFormulaBarVisible(),
      backstageOpen: handle.getBackstageOpen(),
    });
  };
  handle = mountLegacyToolbar(host, legacy, {
    dynamicDropdowns: true,
    hooks: options.hooks,
    ribbonTabs: options.ribbonTabs,
    activeTab: options.activeTab,
    ribbonDisplayMode: options.displayMode,
    formulaBarVisible: options.formulaBarVisible,
    theme: themeOf(),
    onTabChange: notify,
    onDisplayModeChange: notify,
    onFormulaBarChange: notify,
    onBackstageOpenChange: notify,
  });
  const mounted = handle;
  // The grid's theme reaches an inline ribbon through the cascade; this host sits outside it.
  const unsubscribeTheme = legacy.store.subscribe(() => {
    const theme = themeOf();
    if (host.dataset.fcTheme === theme) return;
    host.dataset.fcTheme = theme;
    mounted.setTheme(theme);
  });
  const toolbar = new ToolbarImpl(handle, legacy, events, () => {
    unsubscribeTheme();
    if (previousTheme === undefined) delete host.dataset.fcTheme;
    else host.dataset.fcTheme = previousTheme;
  });
  if (signal?.aborted) {
    toolbar.dispose();
    signal.throwIfAborted();
  }
  return toolbar;
}
