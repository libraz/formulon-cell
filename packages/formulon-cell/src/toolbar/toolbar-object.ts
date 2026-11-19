import { Emitter } from '../api/events.js';
import { type Disposable, Lifecycle } from '../api/lifecycle.js';
import { legacyInstanceOf, type SpreadsheetInstance } from '../api/spreadsheet/instance.js';
import { mountToolbar as mountLegacyToolbar } from '../mount/toolbar.js';
import type {
  MountToolbarOptions as LegacyToolbarOptions,
  ToolbarInstance,
} from '../mount/toolbar-types.js';
import type { SpreadsheetInstance as LegacyInstance } from '../mount/types.js';
import type { RibbonHooks } from './ribbon/apply-ribbon-command.js';
import type { RibbonDisplayMode } from './ribbon/render-ribbon.js';
import type { RibbonTab } from './ribbon-model.js';
import {
  type InstanceRibbon,
  type RibbonEvents,
  RibbonView,
  ribbonStateOf,
} from './ribbon-view.js';

export type { InstanceRibbon, ToolbarPatch, ToolbarState } from './ribbon-view.js';

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

/** A ribbon mounted on its own with `mountToolbar`. */
export interface Toolbar extends Disposable, InstanceRibbon {}

class ToolbarImpl extends RibbonView implements Toolbar {
  readonly #life: Lifecycle;

  constructor(
    handle: ToolbarInstance,
    legacy: LegacyInstance,
    events: Emitter<RibbonEvents>,
    release: () => void,
    life = new Lifecycle('Toolbar'),
  ) {
    super(handle, legacy, events, () => life.assertLive());
    this.#life = life;
    life.onDispose(() => {
      events.clear();
      release();
      handle.dispose();
    });
  }

  get disposed(): boolean {
    return this.#life.disposed;
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
  const events = new Emitter<RibbonEvents>();
  let handle: ToolbarInstance | null = null;
  const notify = (): void => {
    if (handle) events.emit('change', ribbonStateOf(handle));
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
