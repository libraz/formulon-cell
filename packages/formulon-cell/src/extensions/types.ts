// Extension contract — the public unit of feature composition.
//
// Heavily inspired by CodeMirror 6's extension model and Tiptap's "Kit"
// pattern: each extension is a small factory that, when run during mount,
// hooks DOM listeners + store subscriptions and returns a Handle the
// instance can later disable / dispose / rebind.
import type { History } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';

/** Reactive strings controller — Phase 3 makes labels live-updatable. */
export interface I18nController {
  readonly locale: string;
  readonly strings: Strings;
  setLocale(locale: string): void;
  extend(locale: string, overlay: import('../i18n/strings.js').DeepPartial<Strings>): void;
  register(locale: string, strings: Strings): void;
  subscribe(fn: (s: Strings) => void): () => void;
}

/** Resolved theme name — extensions can change theme via this controller. */
export type ThemeName = 'paper' | 'ink' | (string & {});

/** Shared services every extension's `setup` is given. */
export interface ExtensionContext {
  readonly host: HTMLElement;
  readonly formulabar: HTMLElement;
  readonly grid: HTMLElement;
  readonly statusbar: HTMLElement;
  readonly canvas: HTMLCanvasElement;
  readonly a11y: HTMLElement;
  readonly store: SpreadsheetStore;
  readonly history: History;
  readonly i18n: I18nController;
  /** Returns the current workbook. Always read through this — `setWorkbook`
   *  swaps the underlying handle and the cached reference would go stale. */
  getWb(): WorkbookHandle;
  /** Re-pull cell data from the engine into the store. Use after any
   *  mutation that bypasses the store (e.g. paste, fill). */
  refreshCells(): void;
  /** Triggers a re-paint without mutating state. */
  invalidate(): void;
  /** Look up another extension's handle by id. Returns undefined if the
   *  extension wasn't loaded — callers must handle absence. */
  resolve<T extends ExtensionHandle = ExtensionHandle>(id: string): T | undefined;
  /** Subscribe to workbook swaps. Extensions that hold a wb reference
   *  should rebind in the callback. */
  onWorkbookChange(fn: (wb: WorkbookHandle) => void): () => void;
}

/** Returned from an extension's `setup`. Domain-specific public methods
 *  (e.g. `open()`, `refresh()`) sit alongside the lifecycle hooks. */
export interface ExtensionHandle {
  /** Tear down DOM listeners + store subscriptions. Idempotent. */
  dispose(): void;
  /** Optional: rebind to a fresh workbook. mount.ts calls this during
   *  `setWorkbook`. Default behavior (omit method) = re-read via `getWb()`
   *  on next event and don't keep a wb reference. */
  rebindWorkbook?(wb: WorkbookHandle): void;
  /** Optional: re-render labels for a new locale. */
  setStrings?(s: Strings): void;
  /** Optional: temporarily disable without disposing. */
  enable?(): void;
  disable?(): void;
  /** Domain methods (e.g. `open`, `close`, `refresh`) sit here. */
  [key: string]: unknown;
}

/** A composable feature unit. Returned from factory functions like
 *  `findReplace()`, `statusBar()`, `keymap.excel`, etc. */
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
 *  for `presets.excel().concat([myCustomFindReplace()])` overrides. */
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
