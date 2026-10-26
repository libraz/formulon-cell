import type { SpreadsheetStore } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Mirror the workbook's reference style, date system and precision into the store. */
export function hydrateCalcSettingsFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  const settings = wb.calcSettings();
  const r1c1 = settings?.r1c1 ?? false;
  const date1904 = settings?.date1904 ?? false;
  const precisionAsDisplayed = settings ? !settings.fullPrecision : false;
  store.setState((s) =>
    s.ui.r1c1 === r1c1 &&
    s.ui.date1904 === date1904 &&
    s.ui.precisionAsDisplayed === precisionAsDisplayed
      ? s
      : { ...s, ui: { ...s.ui, r1c1, date1904, precisionAsDisplayed } },
  );
}

/** Write R1C1 toggles made anywhere in the UI through to `<calcPr refMode>`,
 *  so the reference style saves with the workbook. */
export function bindCalcSettingsStore(wb: WorkbookHandle, store: SpreadsheetStore): () => void {
  let last = store.getState().ui.r1c1;
  return store.subscribe((state) => {
    if (state.ui.r1c1 === last) return;
    last = state.ui.r1c1;
    wb.setCalcSetting('r1c1', last);
  });
}
