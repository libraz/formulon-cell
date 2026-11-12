import type {
  I18n,
  SelectionState,
  SpreadsheetInstance,
  SpreadsheetInstanceEvents,
} from '@libraz/formulon-cell';
import {
  isLiveSpreadsheet,
  type SpreadsheetSubscriber,
  subscribeToLocale,
  subscribeToSelection,
  subscribeToSpreadsheet,
} from '@libraz/formulon-cell/adapter';
import {
  type MaybeRefOrGetter,
  type ShallowRef,
  shallowRef,
  toValue,
  triggerRef,
  watchEffect,
} from 'vue';

type Source = MaybeRefOrGetter<SpreadsheetInstance | null>;

/** A ref holding `read(spreadsheet)`, re-read whenever `subscribe` notifies; `null` while the
 *  spreadsheet is `null` or disposed. Follows `source` and unsubscribes with the scope. */
const useDerived = <T>(
  source: Source,
  subscribe: SpreadsheetSubscriber,
  read: (spreadsheet: SpreadsheetInstance) => T,
): Readonly<ShallowRef<T | null>> => {
  const out = shallowRef<T | null>(null);
  watchEffect((onCleanup) => {
    const spreadsheet = toValue(source);
    if (!isLiveSpreadsheet(spreadsheet)) {
      out.value = null;
      return;
    }
    const refresh = (): void => {
      out.value = isLiveSpreadsheet(spreadsheet) ? read(spreadsheet) : null;
      // A selector may return the same object after a change; listeners still need the signal.
      triggerRef(out);
    };
    refresh();
    onCleanup(subscribe(spreadsheet, refresh));
  });
  return out;
};

/** Derives a value from the spreadsheet, re-run after every change, selection move, locale change
 *  and `update`. `null` while the spreadsheet is `null` or disposed. */
export const useSpreadsheet = <T>(
  spreadsheet: Source,
  selector: (spreadsheet: SpreadsheetInstance) => T,
): Readonly<ShallowRef<T | null>> => useDerived(spreadsheet, subscribeToSpreadsheet, selector);

/** The current selection; `null` while the spreadsheet is `null` or disposed. */
export const useSelection = (spreadsheet: Source): Readonly<ShallowRef<SelectionState | null>> =>
  useDerived(spreadsheet, subscribeToSelection, (s) => s.selection.get());

/** The active locale and strings; `null` while the spreadsheet is `null` or disposed. */
export const useI18n = (
  spreadsheet: Source,
): Readonly<
  ShallowRef<{ readonly locale: I18n['locale']; readonly strings: I18n['strings'] } | null>
> =>
  useDerived(spreadsheet, subscribeToLocale, (s) => ({
    locale: s.i18n.locale,
    strings: s.i18n.strings,
  }));

/** Subscribes `handler` to an instance event while the spreadsheet is set; follows `spreadsheet`
 *  and unsubscribes with the scope. */
export const useSpreadsheetEvent = <K extends keyof SpreadsheetInstanceEvents>(
  spreadsheet: Source,
  event: K,
  handler: (event: SpreadsheetInstanceEvents[K]) => void,
): void => {
  watchEffect((onCleanup) => {
    const s = toValue(spreadsheet);
    if (!isLiveSpreadsheet(s)) return;
    onCleanup(s.on(event, (e) => handler(e)));
  });
};
