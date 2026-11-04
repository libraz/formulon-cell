import type { CellFormatStrings } from '../cell-format/types.js';
import type { RibbonStrings } from '../ribbon/types.js';

/** Labels of the Mac Zoom dialog, which offers presets instead of a single field. */
export interface MacZoomDialogStrings {
  /** Caption above the magnification choices. */
  caption: string;
  fitSelection: string;
  /** Radio label in front of the custom percentage field. */
  custom: string;
}

/** A locale whose Mac wording is not measured reuses its shared zoom labels. */
export const macZoomDialogFrom = (
  ribbon: RibbonStrings['ribbon'],
  formatDialog: CellFormatStrings['formatDialog'],
): MacZoomDialogStrings => ({
  caption: ribbon.zoomDialogPercent,
  fitSelection: ribbon.zoomSelection,
  custom: formatDialog.catCustom,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macZoomDialogJa: MacZoomDialogStrings = {
  caption: '倍率:',
  fitSelection: '選択範囲をズーム',
  custom: 'カスタム:',
};
