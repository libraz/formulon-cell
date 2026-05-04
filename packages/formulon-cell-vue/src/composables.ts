import type {
  SpreadsheetEventHandler,
  SpreadsheetEventName,
  SpreadsheetInstance,
} from '@libraz/formulon-cell';
import { onScopeDispose, type Ref, ref, watchEffect } from 'vue';

/** Composable: track the live selection from a `ref<SpreadsheetInstance | null>`. */
export const useSelection = (
  instance: Ref<SpreadsheetInstance | null>,
): Ref<ReturnType<SpreadsheetInstance['store']['getState']>['selection']> => {
  const sel = ref({
    active: { sheet: 0, row: 0, col: 0 },
    anchor: { sheet: 0, row: 0, col: 0 },
    range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
  }) as Ref<ReturnType<SpreadsheetInstance['store']['getState']>['selection']>;

  let off: (() => void) | null = null;
  watchEffect(() => {
    off?.();
    const inst = instance.value;
    if (!inst) return;
    sel.value = inst.store.getState().selection;
    off = inst.store.subscribe(() => {
      sel.value = inst.store.getState().selection;
    });
  });
  onScopeDispose(() => off?.());
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
  const out = ref(fallback) as Ref<T>;
  let off: (() => void) | null = null;
  watchEffect(() => {
    off?.();
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
  onScopeDispose(() => off?.());
  return out;
};

/** Composable: track the live i18n locale + strings. */
export const useI18n = (
  instance: Ref<SpreadsheetInstance | null>,
): {
  locale: Ref<string>;
  strings: Ref<SpreadsheetInstance['i18n']['strings']>;
} => {
  const locale = ref<string>('ja');
  const strings = ref({}) as Ref<SpreadsheetInstance['i18n']['strings']>;

  let off: (() => void) | null = null;
  watchEffect(() => {
    off?.();
    const inst = instance.value;
    if (!inst) return;
    locale.value = inst.i18n.locale;
    strings.value = inst.i18n.strings;
    off = inst.i18n.subscribe((next) => {
      locale.value = inst.i18n.locale;
      strings.value = next;
    });
  });
  onScopeDispose(() => off?.());
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
