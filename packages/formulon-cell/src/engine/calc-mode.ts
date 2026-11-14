/** `<calcPr calcMode>` codes, as `WorkbookHandle.calcMode()` reports them. */
export type CalcModeCode = 0 | 1 | 2;

export const CALC_MODE_MANUAL = 1;
/** Automatic except for data tables. */
export const CALC_MODE_AUTO_NO_TABLE = 2;

/** The mode a status-bar calc-mode toggle steps to: Auto → Manual → Auto-except-tables → Auto. */
export const nextCalcMode = (mode: number): CalcModeCode => ((mode + 1) % 3) as CalcModeCode;

/** Whether `mode` can leave formula values stale: Manual, or Automatic except for data tables. */
export const calcModeLeavesStaleValues = (mode: CalcModeCode | null): boolean =>
  mode === CALC_MODE_MANUAL || mode === CALC_MODE_AUTO_NO_TABLE;

/** The Compatibility Version codes the engine setter takes. 3 is stored as given and evaluates
 *  as 2; a loaded stamp above 3 reads back as its own number. */
export type CompatibilityVersionCode = 1 | 2 | 3;
