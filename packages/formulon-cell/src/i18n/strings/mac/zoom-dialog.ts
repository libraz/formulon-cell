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

/** Mac app wording from the Excel bundle; not answered by it: custom. */
export const macZoomDialogDe: Partial<MacZoomDialogStrings> = {
  caption: 'Zoommodus:',
  fitSelection: 'An Markierung anpassen',
};

/** Mac app wording from the Excel bundle; not answered by it: custom. */
export const macZoomDialogFr: Partial<MacZoomDialogStrings> = {
  caption: 'Facteur d’agrandissement :',
  fitSelection: 'Ajuster à la sélection',
};

/** Mac app wording from the Excel bundle; not answered by it: custom. */
export const macZoomDialogKo: Partial<MacZoomDialogStrings> = {
  caption: '배율:',
  fitSelection: '선택 영역에 맞춤',
};

/** Mac app wording from the Excel bundle; not answered by it: custom. */
export const macZoomDialogTh: Partial<MacZoomDialogStrings> = {
  caption: 'อัตราส่วน:',
  fitSelection: 'จัดส่วนที่เลือกให้พอดีกับหน้าจอ',
};

/** Mac app wording from the Excel bundle; not answered by it: custom. */
export const macZoomDialogZh: Partial<MacZoomDialogStrings> = {
  caption: '缩放:',
  fitSelection: '适应所选内容',
};

/** Mac app wording from the Excel bundle. */
export const macZoomDialogEn: Partial<MacZoomDialogStrings> = {
  fitSelection: 'Fit selection',
};
