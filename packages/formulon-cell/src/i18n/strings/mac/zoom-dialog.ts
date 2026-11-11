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

/** Each locale starts from its shared zoom labels and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: fitSelection, custom. */
export const macZoomDialogDe: Partial<MacZoomDialogStrings> = {
  caption: 'Zoommodus:',
};

/** Mac app wording from the Excel bundle; not answered by it: fitSelection, custom. */
export const macZoomDialogFr: Partial<MacZoomDialogStrings> = {
  caption: 'Facteur d’agrandissement :',
};

/** Mac app wording from the Excel bundle; not answered by it: fitSelection, custom. */
export const macZoomDialogKo: Partial<MacZoomDialogStrings> = {
  caption: '배율:',
};

/** Mac app wording from the Excel bundle; not answered by it: fitSelection, custom. */
export const macZoomDialogTh: Partial<MacZoomDialogStrings> = {
  caption: 'อัตราส่วน:',
};

/** Mac app wording from the Excel bundle; not answered by it: fitSelection, custom. */
export const macZoomDialogZh: Partial<MacZoomDialogStrings> = {
  caption: '缩放:',
};
