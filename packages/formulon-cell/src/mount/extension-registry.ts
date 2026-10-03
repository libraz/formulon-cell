import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  dedupeById,
  type Extension,
  type ExtensionContext,
  type ExtensionHandle,
  type ExtensionInput,
  flattenExtensions,
  sortByPriority,
} from '../extensions/index.js';
import type { Strings } from '../i18n/strings.js';
import { type EngineBinding, WB_REGISTRY_IDS } from './engine-binding.js';

export interface ExtensionRegistryDeps {
  /** Extension context minus the lookups the registry itself provides. */
  context: Omit<ExtensionContext, 'resolve' | 'onWorkbookChange'>;
}

export interface ExtensionRegistry {
  /** Built-in feature handles, keyed by feature id. */
  featureRegistry: Map<string, ExtensionHandle>;
  /** Combined built-in + user view; mutated in place so the identity is stable. */
  featuresView: Record<string, ExtensionHandle | undefined>;
  wrapHandle: (raw: unknown, detach: () => void) => ExtensionHandle;
  /** Mirror the engine binding's handles into `featureRegistry`. */
  syncBindingFeatures: (current: EngineBinding) => void;
  refreshFeaturesView: () => void;
  getUserHandle: (id: string) => ExtensionHandle | undefined;
  /** Mount `input` (flattened, deduped, priority-sorted) without refreshing the view. */
  mountAll: (input: readonly ExtensionInput[]) => void;
  use: (input: ExtensionInput) => void;
  remove: (id: string) => boolean;
  setExtensions: (next: ExtensionInput[] | undefined) => void;
  setUserStrings: (next: Strings) => void;
  /** Tell user extensions and workbook listeners that `wb` replaced the previous workbook. */
  notifyWorkbookChange: (wb: WorkbookHandle) => void;
  /** Dispose every user extension. */
  dispose: () => void;
}

export function createExtensionRegistry(deps: ExtensionRegistryDeps): ExtensionRegistry {
  const featureRegistry = new Map<string, ExtensionHandle>();
  const wrapHandle = (raw: unknown, detach: () => void): ExtensionHandle => {
    const h = (
      raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {}
    ) as ExtensionHandle;
    h.dispose = detach;
    return h;
  };

  const syncBindingFeatures = (current: EngineBinding): void => {
    for (const id of WB_REGISTRY_IDS) featureRegistry.delete(id);
    if (current.clipboardH) {
      featureRegistry.set(
        'clipboard',
        wrapHandle(current.clipboardH, () => current.clipboardH?.detach()),
      );
    }
    if (current.pasteSpecialDialog) {
      featureRegistry.set(
        'pasteSpecial',
        wrapHandle(current.pasteSpecialDialog, () => current.pasteSpecialDialog?.detach()),
      );
    }
    if (current.quickAnalysis) {
      featureRegistry.set(
        'quickAnalysis',
        wrapHandle(current.quickAnalysis, () => current.quickAnalysis?.detach()),
      );
    }
    if (current.contextMenu) featureRegistry.set('contextMenu', current.contextMenu);
    if (current.findReplace) {
      featureRegistry.set(
        'findReplace',
        wrapHandle(current.findReplace, () => current.findReplace?.detach()),
      );
    }
    if (current.validation) {
      featureRegistry.set(
        'validation',
        wrapHandle(current.validation, () => current.validation?.detach()),
      );
    }
  };

  // User extensions — additive on top of built-ins. Run after built-ins
  // and the engine binding so they can read other features via
  // `ctx.resolve()`.
  const userHandles = new Map<string, ExtensionHandle>();
  const wbListeners = new Set<(next: WorkbookHandle) => void>();
  const ctx: ExtensionContext = {
    ...deps.context,
    resolve: <T extends ExtensionHandle = ExtensionHandle>(id: string): T | undefined =>
      (featureRegistry.get(id) ?? userHandles.get(id)) as T | undefined,
    onWorkbookChange: (fn) => {
      wbListeners.add(fn);
      return () => {
        wbListeners.delete(fn);
      };
    },
  };

  const mountExtension = (ext: Extension): void => {
    if (userHandles.has(ext.id) || featureRegistry.has(ext.id)) {
      // last-wins via remove + re-add; users explicitly opt in
      userHandles.get(ext.id)?.dispose();
      userHandles.delete(ext.id);
    }
    const handle = ext.setup(ctx);
    if (handle) userHandles.set(ext.id, handle);
  };
  const mountAll = (input: readonly ExtensionInput[]): void => {
    const sorted = sortByPriority(dedupeById(flattenExtensions(input)));
    for (const ext of sorted) mountExtension(ext);
  };
  // Combined view exposed on `instance.features` — built-ins + user.
  const featuresView: Record<string, ExtensionHandle | undefined> = {};
  const refreshFeaturesView = (): void => {
    for (const k of Object.keys(featuresView)) delete featuresView[k];
    for (const [k, v] of featureRegistry) featuresView[k] = v;
    for (const [k, v] of userHandles) featuresView[k] = v;
  };
  const disposeUserHandles = (): void => {
    for (const handle of userHandles.values()) handle.dispose();
    userHandles.clear();
  };

  return {
    featureRegistry,
    featuresView,
    wrapHandle,
    syncBindingFeatures,
    refreshFeaturesView,
    getUserHandle: (id) => userHandles.get(id),
    mountAll,
    use(input) {
      mountAll([input]);
      refreshFeaturesView();
    },
    remove(id) {
      const handle = userHandles.get(id);
      if (!handle) return false;
      handle.dispose();
      userHandles.delete(id);
      refreshFeaturesView();
      return true;
    },
    setExtensions(next) {
      // Dispose all currently-mounted user extensions, then re-mount the
      // new list. Built-ins are untouched — use `setFeatures` for those.
      disposeUserHandles();
      if (next?.length) mountAll(next);
      refreshFeaturesView();
    },
    setUserStrings(next) {
      for (const handle of userHandles.values()) handle.setStrings?.(next);
    },
    notifyWorkbookChange(wb) {
      // Notify user extensions so they can rebind their wb references.
      for (const handle of [...userHandles.values()]) {
        try {
          handle.rebindWorkbook?.(wb);
        } catch (error) {
          console.warn('formulon-cell: extension workbook hook failed', error);
        }
      }
      for (const fn of [...wbListeners]) {
        try {
          fn(wb);
        } catch (error) {
          console.warn('formulon-cell: workbook hook failed', error);
        }
      }
    },
    dispose: disposeUserHandles,
  };
}
