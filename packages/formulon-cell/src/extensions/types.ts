// Extension contract — the public unit of feature composition.
//
// Heavily inspired by CodeMirror 6's extension model and Tiptap's "Kit"
// pattern: each extension is a small factory that, when run during mount,
// hooks DOM listeners + store subscriptions and returns a Handle the
// instance can later disable / dispose / rebind.

import type { Workbook } from '../api/workbook.js';
import type { History } from '../commands/history.js';
import type { HistoryView } from '../commands/history-view.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { I18nController } from '../i18n/controller.js';
import type { I18n } from '../i18n/public-view.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';

/** Built-in theme ids; any other string selects a host-defined theme. */
export type ThemeName = 'paper' | 'ink' | 'contrast' | (string & {});

/** Shared services every extension's `setup` is given. */
export interface ExtensionContext {
  readonly host: HTMLElement;
  readonly formulabar: HTMLElement;
  readonly viewbar: HTMLElement;
  readonly grid: HTMLElement;
  readonly statusbar: HTMLElement;
  readonly canvas: HTMLCanvasElement;
  /** Polite live region for announcements. Writing text here is announced by
   *  assistive tech; the structured cell mirror lives inside the grid. */
  readonly a11y: HTMLElement;
  /** The instance's UI language and strings. */
  readonly i18n: I18n;
  /** The current workbook. Read it through the context each time: swapping the workbook
   *  replaces it, so a cached reference goes stale. */
  readonly workbook: Workbook;
  /** The instance's undo history and workbook-swap event. */
  readonly spreadsheet: ExtensionSpreadsheet;
  /** Re-pull cell data from the engine into the store. Use after any
   *  mutation that bypasses the store (e.g. paste, fill). */
  refreshCells(): void;
  /** Triggers a re-paint without mutating state. */
  invalidate(): void;
  /** Look up another extension's handle by id. Returns undefined if the
   *  extension wasn't loaded — callers must handle absence. */
  resolve<T extends ExtensionHandle = ExtensionHandle>(id: string): T | undefined;
}

/** The parts of the instance an extension can use. */
export interface ExtensionSpreadsheet {
  readonly history: HistoryView;
  /** Subscribes to workbook swaps; an extension holding a workbook reference rebinds in the
   *  listener. The returned disposer is idempotent. */
  on(name: 'workbook', listener: (event: { workbook: Workbook }) => void): () => void;
}

/** What the built-in features are given: the public context with the engine handle and the
 *  controller in place of the facade and the read-only view, plus the store and the raw history. */
export interface BuiltInExtensionContext extends Omit<ExtensionContext, 'workbook' | 'i18n'> {
  readonly workbook: WorkbookHandle;
  readonly i18n: I18nController;
  readonly store: SpreadsheetStore;
  readonly history: History;
  /** Subscribe to workbook swaps with the engine handle. */
  onWorkbookChange(fn: (wb: WorkbookHandle) => void): () => void;
}

/** Returned from an extension's `setup`. Domain-specific public methods
 *  (e.g. `open()`, `refresh()`) sit alongside the lifecycle hooks. */
export interface ExtensionHandle {
  /** Tear down DOM listeners + store subscriptions. Idempotent. */
  dispose(): void;
  /** Optional: re-render labels for a new locale. */
  setStrings?(s: Strings): void;
  /** Optional: temporarily disable without disposing. */
  enable?(): void;
  disable?(): void;
  /** Domain methods (e.g. `open`, `close`, `refresh`) sit here. */
  [key: string]: unknown;
}

/** The handle of a built-in feature; the registry rebinds it to the engine handle on a swap. */
export interface BuiltInExtensionHandle extends ExtensionHandle {
  rebindWorkbook?(wb: WorkbookHandle): void;
}

/** A composable feature unit. Returned from factory functions like
 *  `findReplace()`, `statusBar()`, keymap helpers, etc. */
export interface Extension {
  /** Stable identifier — used as the registry key + `instance.features`
   *  property name. Each id should appear at most once in a mount. */
  readonly id: string;
  /** Lower runs earlier. Defaults: 0=core/i18n/theme, 10=DOM chrome,
   *  50=engine listeners, 80=cross-cutting (context menu, keymap). */
  readonly priority?: number;
  /** Called once during mount. Returns a handle, or void if there's
   *  nothing to dispose. Extensions can be passed nested in arrays;
   *  the registry flattens them. */
  setup(ctx: ExtensionContext): ExtensionHandle | void;
}

/** A built-in feature unit: `Extension` with the engine-level context and handle. */
export interface BuiltInExtension {
  readonly id: string;
  readonly priority?: number;
  setup(ctx: BuiltInExtensionContext): BuiltInExtensionHandle | void;
}

/** Allow nested arrays so presets compose naturally. */
export type ExtensionInput = Extension | readonly ExtensionInput[];

export const flattenExtensions = (input: readonly ExtensionInput[]): Extension[] => {
  const out: Extension[] = [];
  const walk = (x: ExtensionInput): void => {
    if (Array.isArray(x)) {
      for (const child of x) walk(child);
    } else {
      out.push(x as Extension);
    }
  };
  for (const x of input) walk(x);
  return out;
};

/** Stable sort by priority (lower = earlier). Ties preserve input order. */
export const sortByPriority = (exts: Extension[]): Extension[] => {
  return exts
    .map((ext, idx) => ({ ext, idx }))
    .sort((a, b) => (a.ext.priority ?? 50) - (b.ext.priority ?? 50) || a.idx - b.idx)
    .map((entry) => entry.ext);
};

/** Drop any extension whose id is later in the list — last-wins. Useful
 *  for `presets.full().concat([myCustomFindReplace()])` overrides. */
export const dedupeById = (exts: Extension[]): Extension[] => {
  const seen = new Set<string>();
  const out: Extension[] = [];
  for (let i = exts.length - 1; i >= 0; i -= 1) {
    const ext = exts[i];
    if (!ext) continue;
    if (seen.has(ext.id)) continue;
    seen.add(ext.id);
    out.unshift(ext);
  }
  return out;
};
