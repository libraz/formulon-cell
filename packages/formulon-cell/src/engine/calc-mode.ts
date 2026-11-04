/** `<calcPr calcMode>` codes, as `WorkbookHandle.calcMode()` reports them. */
export type CalcModeCode = 0 | 1 | 2;

export const CALC_MODE_MANUAL = 1;
/** Automatic except for data tables. */
export const CALC_MODE_AUTO_NO_TABLE = 2;

/** The mode a status-bar calc-mode toggle steps to: Auto → Manual → Auto-except-tables → Auto. */
export const nextCalcMode = (mode: number): CalcModeCode => ((mode + 1) % 3) as CalcModeCode;
