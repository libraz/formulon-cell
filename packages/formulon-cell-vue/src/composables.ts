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

/** Composable: track the live selection from a `ref<SpreadsheetInstance | null>`. */
export const useSelection = (instance: Ref<SpreadsheetInstance | null>): Ref<Selection> =>
  useSpreadsheet(instance, (state) => state.selection, ADAPTER_FALLBACK_SELECTION);

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
 *  `localeChange`, `themeChange`, `recalc`). The subscription follows the
 *  instance ref and is torn down with the scope. */
export const useSpreadsheetEvent = <K extends SpreadsheetEventName>(
  instance: Ref<SpreadsheetInstance | null>,
  event: K,
  handler: SpreadsheetEventHandler<K>,
): void => {
  watchEffect((onCleanup) => {
    const inst = instance.value;
    if (!inst) return;
    onCleanup(inst.on(event, handler));
  });
};
