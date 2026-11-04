import {
  ADAPTER_FALLBACK_LOCALE,
  ADAPTER_FALLBACK_SELECTION,
  type SpreadsheetEventHandler,
  type SpreadsheetEventName,
  type SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { onScopeDispose, type Ref, shallowRef, watchEffect } from 'vue';

type Selection = ReturnType<SpreadsheetInstance['store']['getState']>['selection'];

const createFallbackStrings = (): SpreadsheetInstance['i18n']['strings'] =>
  ({}) as unknown as SpreadsheetInstance['i18n']['strings'];

/** Composable: track the live selection from a `ref<SpreadsheetInstance | null>`. */
export const useSelection = (instance: Ref<SpreadsheetInstance | null>): Ref<Selection> => {
  const sel = shallowRef<Selection>(ADAPTER_FALLBACK_SELECTION);

  let off: (() => void) | null = null;
  watchEffect(() => {
    const previousOff = off;
    off = null;
    previousOff?.();
    const inst = instance.value;
    if (!inst) {
      sel.value = ADAPTER_FALLBACK_SELECTION;
      return;
    }
    sel.value = inst.store.getState().selection;
    off = inst.store.subscribe(() => {
      sel.value = inst.store.getState().selection;
    });
  });
  onScopeDispose(() => {
    const currentOff = off;
    off = null;
    currentOff?.();
  });
  return sel;
};

/** Composable: derive a value from the spreadsheet store. The selector is
 *  re-run on every store change. Falls back to `fallback` while the
 *  instance is null (mount-pending / disposed). */
export const useSpreadsheet = <T>(
  instance: Ref<SpreadsheetInstance | null>,
  selector: (state: ReturnType<SpreadsheetInstance['store']['getState']>) => T,
  fallback: T,
): Ref<T> => {
  const out = shallowRef<T>(fallback);
  let off: (() => void) | null = null;
  watchEffect(() => {
    const previousOff = off;
    off = null;
    previousOff?.();
    const inst = instance.value;
    if (!inst) {
      out.value = fallback;
      return;
    }
    out.value = selector(inst.store.getState());
    off = inst.store.subscribe(() => {
      out.value = selector(inst.store.getState());
    });
  });
  onScopeDispose(() => {
    const currentOff = off;
    off = null;
    currentOff?.();
  });
  return out;
};

/** Composable: track the live i18n locale + strings. */
export const useI18n = (
  instance: Ref<SpreadsheetInstance | null>,
): {
  locale: Ref<string>;
  strings: Ref<SpreadsheetInstance['i18n']['strings']>;
} => {
  const locale = shallowRef<string>(ADAPTER_FALLBACK_LOCALE);
  const strings = shallowRef(createFallbackStrings());

  let off: (() => void) | null = null;
  watchEffect(() => {
    const previousOff = off;
    off = null;
    previousOff?.();
    const inst = instance.value;
    if (!inst) {
      locale.value = ADAPTER_FALLBACK_LOCALE;
      strings.value = createFallbackStrings();
      return;
    }
    locale.value = inst.i18n.locale;
    strings.value = inst.i18n.strings;
    off = inst.i18n.subscribe((next) => {
      locale.value = inst.i18n.locale;
      strings.value = next;
    });
  });
  onScopeDispose(() => {
    const currentOff = off;
    off = null;
    currentOff?.();
  });
  return { locale, strings };
};

/** Composable: subscribe to a named lifecycle event on the spreadsheet
 *  instance (`cellChange`, `selectionChange`, `workbookChange`,
 *  `localeChange`, `themeChange`, `recalc`). The handler reference can
 *  change between renders without re-subscribing — the latest one is
 *  invoked on each event. */
export const useSpreadsheetEvent = <K extends SpreadsheetEventName>(
  instance: Ref<SpreadsheetInstance | null>,
  event: K,
  handler: SpreadsheetEventHandler<K>,
): void => {
  let currentHandler = handler;
  watchEffect((onCleanup) => {
    currentHandler = handler;
    const inst = instance.value;
    if (!inst) return;
    const off = inst.on(event, (e) => currentHandler(e));
    onCleanup(off);
  });
};
