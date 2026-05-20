// Alignment tab DOM for the Format Cells dialog. Extracted from
// `format-dialog-view.ts`; the parent file binds the returned refs to event
// handlers and the central state machine.

import type { Strings } from '../../i18n/strings.js';
import type { CellAlign, CellVAlign, TextDirection } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { createDialogToggleButton } from '../dialog-shell.js';
import { makeCheckbox } from '../format-dialog-dom.js';

export interface AlignTabRefs {
  hAlignRadios: Map<'default' | CellAlign, HTMLInputElement>;
  hAlignSelect: HTMLSelectElement;
  vAlignRadios: Map<'default' | CellVAlign, HTMLInputElement>;
  vAlignSelect: HTMLSelectElement;
  wrapCk: ReturnType<typeof makeCheckbox>;
  shrinkCk: ReturnType<typeof makeCheckbox>;
  mergeCk: ReturnType<typeof makeCheckbox>;
  indentInput: HTMLInputElement;
  textDirectionSelect: HTMLSelectElement;
  rotationInput: HTMLInputElement;
  alignPreviewDial: HTMLDivElement;
  alignPreviewDialDots: HTMLButtonElement[];
  alignPreviewDialPointer: HTMLSpanElement;
  alignPreviewDialText: HTMLSpanElement;
}

export function createAlignTab(panel: HTMLDivElement, t: Strings['formatDialog']): AlignTabRefs {
  // Horizontal
  const hAlignLegend = document.createElement('div');
  hAlignLegend.textContent = t.horizontalAlign;
  panel.appendChild(hAlignLegend);
  const hAlignFieldset = document.createElement('div');
  hAlignFieldset.className = 'fc-fmtdlg__choice-grid';
  hAlignFieldset.setAttribute('role', 'radiogroup');
  hAlignFieldset.setAttribute('aria-label', t.horizontalAlign);
  panel.appendChild(hAlignFieldset);

  const hAlignName = `fc-fmtdlg-halign-${Math.random().toString(36).slice(2, 8)}`;
  const hAlignDefs: { id: 'default' | CellAlign; label: string }[] = [
    { id: 'default', label: t.alignDefault },
    { id: 'left', label: t.alignLeft },
    { id: 'center', label: t.alignCenter },
    { id: 'right', label: t.alignRight },
    { id: 'fill', label: t.alignFill },
    { id: 'justify', label: t.alignJustify },
    { id: 'centerContinuous', label: t.alignCenterAcrossSelection },
    { id: 'distributed', label: t.alignDistributed },
  ];
  const hAlignRadios = new Map<'default' | CellAlign, HTMLInputElement>();
  for (const a of hAlignDefs) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__radio';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = hAlignName;
    radio.value = a.id;
    const txt = document.createElement('span');
    txt.textContent = a.label;
    wrap.append(radio, txt);
    hAlignFieldset.appendChild(wrap);
    hAlignRadios.set(a.id, radio);
  }

  const hAlignSelectRow = document.createElement('label');
  hAlignSelectRow.className = 'fc-fmtdlg__row fc-fmtdlg__align-select-row';
  const hAlignSelectLabel = document.createElement('span');
  hAlignSelectLabel.textContent = t.horizontalAlign;
  const hAlignSelect = createDialogSelect(
    hAlignDefs.map((a) => ({ value: a.id, label: a.label })),
    'default',
    { ariaLabel: t.horizontalAlign, className: '' },
  );
  hAlignSelect.dataset.fcSelect = 'align';
  hAlignSelectRow.append(hAlignSelectLabel, hAlignSelect);
  panel.appendChild(hAlignSelectRow);

  // Vertical
  const vAlignLegend = document.createElement('div');
  vAlignLegend.textContent = t.verticalAlign;
  panel.appendChild(vAlignLegend);
  const vAlignFieldset = document.createElement('div');
  vAlignFieldset.className = 'fc-fmtdlg__choice-grid';
  vAlignFieldset.setAttribute('role', 'radiogroup');
  vAlignFieldset.setAttribute('aria-label', t.verticalAlign);
  panel.appendChild(vAlignFieldset);

  const vAlignName = `fc-fmtdlg-valign-${Math.random().toString(36).slice(2, 8)}`;
  const vAlignDefs: { id: 'default' | CellVAlign; label: string }[] = [
    { id: 'default', label: t.alignDefault },
    { id: 'top', label: t.vAlignTop },
    { id: 'middle', label: t.vAlignMiddle },
    { id: 'bottom', label: t.vAlignBottom },
    { id: 'justify', label: t.vAlignJustify },
    { id: 'distributed', label: t.vAlignDistributed },
  ];
  const vAlignRadios = new Map<'default' | CellVAlign, HTMLInputElement>();
  for (const a of vAlignDefs) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__radio';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = vAlignName;
    radio.value = a.id;
    const txt = document.createElement('span');
    txt.textContent = a.label;
    wrap.append(radio, txt);
    vAlignFieldset.appendChild(wrap);
    vAlignRadios.set(a.id, radio);
  }

  const vAlignSelectRow = document.createElement('label');
  vAlignSelectRow.className = 'fc-fmtdlg__row fc-fmtdlg__align-select-row';
  const vAlignSelectLabel = document.createElement('span');
  vAlignSelectLabel.textContent = t.verticalAlign;
  const vAlignSelect = createDialogSelect(
    vAlignDefs.map((a) => ({ value: a.id, label: a.label })),
    'default',
    { ariaLabel: t.verticalAlign, className: '' },
  );
  vAlignSelect.dataset.fcSelect = 'vAlign';
  vAlignSelectRow.append(vAlignSelectLabel, vAlignSelect);
  panel.appendChild(vAlignSelectRow);

  // Wrap / Indent / Rotation
  const wrapRow = document.createElement('div');
  wrapRow.className = 'fc-fmtdlg__choice-grid';
  panel.appendChild(wrapRow);
  const wrapCk = makeCheckbox(t.wrap);
  wrapCk.input.dataset.fcCheck = 'wrap';
  wrapRow.append(wrapCk.wrap);

  const indentRow = document.createElement('label');
  indentRow.className = 'fc-fmtdlg__row';
  const indentLabel = document.createElement('span');
  indentLabel.textContent = t.indent;
  const indentInput = document.createElement('input');
  indentInput.type = 'number';
  indentInput.setAttribute('aria-label', t.indent);
  indentInput.min = '0';
  indentInput.max = '15';
  indentInput.step = '1';
  indentRow.append(indentLabel, indentInput);
  panel.appendChild(indentRow);

  const textDirectionRow = document.createElement('label');
  textDirectionRow.className = 'fc-fmtdlg__row';
  const textDirectionLabel = document.createElement('span');
  textDirectionLabel.textContent = t.textDirection;
  const directionDefs: Array<{ id: TextDirection; label: string }> = [
    { id: 'context', label: t.directionContext },
    { id: 'ltr', label: t.directionLeftToRight },
    { id: 'rtl', label: t.directionRightToLeft },
  ];
  const textDirectionSelect = createDialogSelect(
    directionDefs.map((direction) => ({ value: direction.id, label: direction.label })),
    'context',
    { ariaLabel: t.textDirection, className: '' },
  );
  textDirectionSelect.dataset.fcSelect = 'textDirection';
  textDirectionRow.append(textDirectionLabel, textDirectionSelect);
  panel.appendChild(textDirectionRow);

  const rotationRow = document.createElement('label');
  rotationRow.className = 'fc-fmtdlg__row fc-fmtdlg__rotation-row';
  const rotationLabel = document.createElement('span');
  rotationLabel.textContent = t.rotation;
  const rotationInput = document.createElement('input');
  rotationInput.type = 'number';
  rotationInput.setAttribute('aria-label', t.rotation);
  rotationInput.min = '-90';
  rotationInput.max = '90';
  rotationInput.step = '1';
  rotationRow.append(rotationLabel, rotationInput);
  panel.appendChild(rotationRow);

  const alignPreview = document.createElement('div');
  alignPreview.className = 'fc-fmtdlg__align-preview';
  const alignPreviewTitle = document.createElement('div');
  alignPreviewTitle.className = 'fc-fmtdlg__align-preview-title';
  alignPreviewTitle.textContent = t.direction;
  const alignPreviewBox = document.createElement('div');
  alignPreviewBox.className = 'fc-fmtdlg__align-preview-box';
  const alignPreviewVertical = document.createElement('div');
  alignPreviewVertical.className = 'fc-fmtdlg__align-preview-vertical';
  alignPreviewVertical.textContent = t.previewText;
  const alignPreviewDial = document.createElement('div');
  alignPreviewDial.className = 'fc-fmtdlg__align-preview-dial';
  alignPreviewDial.setAttribute('role', 'group');
  alignPreviewDial.setAttribute('aria-label', t.rotation);
  const alignPreviewDialArc = document.createElement('span');
  alignPreviewDialArc.className = 'fc-fmtdlg__align-preview-arc';
  alignPreviewDialArc.setAttribute('aria-hidden', 'true');
  alignPreviewDial.appendChild(alignPreviewDialArc);
  const ANGLE_STEPS = [90, 75, 60, 45, 30, 15, 0, -15, -30, -45, -60, -75, -90];
  const alignPreviewDialDots: HTMLButtonElement[] = [];
  for (const angle of ANGLE_STEPS) {
    const dot = createDialogToggleButton({
      label: `${angle}°`,
      baseClass: 'fc-fmtdlg__align-preview-dot',
      title: `${angle}°`,
      datasetKey: 'fcAngle',
      value: String(angle),
    });
    const rad = (angle * Math.PI) / 180;
    const cx = 12;
    const cy = 66;
    const radius = 56;
    const px = cx + radius * Math.cos(rad);
    const py = cy - radius * Math.sin(rad);
    dot.style.left = `${px}px`;
    dot.style.top = `${py}px`;
    alignPreviewDial.appendChild(dot);
    alignPreviewDialDots.push(dot);
  }
  const alignPreviewDialPointer = document.createElement('span');
  alignPreviewDialPointer.className = 'fc-fmtdlg__align-preview-pointer';
  alignPreviewDialPointer.setAttribute('aria-hidden', 'true');
  alignPreviewDial.appendChild(alignPreviewDialPointer);
  const alignPreviewDialText = document.createElement('span');
  alignPreviewDialText.className = 'fc-fmtdlg__align-preview-text';
  alignPreviewDialText.textContent = t.previewText;
  alignPreviewDial.appendChild(alignPreviewDialText);
  alignPreviewBox.append(alignPreviewVertical, alignPreviewDial);
  const alignPreviewDegree = document.createElement('label');
  alignPreviewDegree.className = 'fc-fmtdlg__align-degree';
  const alignPreviewDegreeLabel = document.createElement('span');
  alignPreviewDegreeLabel.textContent = t.rotation.replace(/\s*\(.*\)$/, '');
  alignPreviewDegree.append(alignPreviewDegreeLabel, rotationInput);
  alignPreview.append(alignPreviewTitle, alignPreviewBox, alignPreviewDegree);
  panel.appendChild(alignPreview);

  const textControl = document.createElement('div');
  textControl.className = 'fc-fmtdlg__text-control';
  const textControlTitle = document.createElement('div');
  textControlTitle.className = 'fc-fmtdlg__text-control-title';
  textControlTitle.textContent = t.textControl;
  const shrinkCk = makeCheckbox(t.shrinkToFit);
  shrinkCk.input.dataset.fcCheck = 'shrinkToFit';
  const mergeCk = makeCheckbox(t.mergeCells);
  mergeCk.input.dataset.fcCheck = 'mergeCells';
  textControl.append(textControlTitle, wrapCk.wrap, shrinkCk.wrap, mergeCk.wrap);
  panel.appendChild(textControl);

  return {
    hAlignRadios,
    hAlignSelect,
    vAlignRadios,
    vAlignSelect,
    wrapCk,
    shrinkCk,
    mergeCk,
    indentInput,
    textDirectionSelect,
    rotationInput,
    alignPreviewDial,
    alignPreviewDialDots,
    alignPreviewDialPointer,
    alignPreviewDialText,
  };
}
