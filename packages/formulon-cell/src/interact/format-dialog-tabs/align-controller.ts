// Alignment tab behaviour for the Format Cells dialog: alignment radios and
// selects, text control checkboxes, indent/direction/rotation and the
// rotation dial. Also owns the merge checkbox's pending edit, which is not
// part of the draft because a merge is a structural change, not a format.

import type { CellAlign, CellVAlign, TextDirection } from '../../store/store.js';
import { projectDisabledReason, projectDisabledState } from '../../toolbar/menu-a11y.js';
import type { AlignTabRefs } from './align.js';
import {
  type FormatTabContext,
  type FormatTabController,
  markMixed,
  markMixedCheck,
} from './controller.js';

export type MergeSelectionState = 'none' | 'merged' | 'mixed';

export interface AlignTabController extends FormatTabController {
  /** Move the dial's marker onto the draft angle and tilt its sample text. */
  syncRotationDial(): void;
  /** Reset the merge checkbox to the selection's merge state. */
  hydrateMerge(selected: MergeSelectionState, disabled: boolean): void;
  /** The merge edit the user made since hydration, if any. */
  mergeAction(): 'merge' | 'unmerge' | null;
}

export function attachAlignTab(ctx: FormatTabContext, refs: AlignTabRefs): AlignTabController {
  const { draft, t, on, touch, isMixed, renderPreview } = ctx;
  const {
    hAlignRadios,
    hAlignSelect,
    vAlignRadios,
    vAlignSelect,
    wrapCk,
    justifyLastLineCk,
    shrinkCk,
    mergeCk,
    indentInput,
    textDirectionSelect,
    rotationInput,
    alignPreviewDial,
    alignPreviewDialDots,
    alignPreviewDialPointer,
    alignPreviewDialText,
  } = refs;
  const merge = { touched: false, initialChecked: false, initialIndeterminate: false };

  const syncRotationDial = (): void => {
    const rotation = draft.rotation;
    for (const dot of alignPreviewDialDots) {
      const angle = Number.parseInt(dot.dataset.fcAngle ?? '0', 10);
      const active = angle === rotation;
      dot.classList.toggle('fc-fmtdlg__align-preview-dot--active', active);
      dot.setAttribute('aria-pressed', active ? 'true' : 'false');
    }
    const rad = (rotation * Math.PI) / 180;
    const cx = 12;
    const cy = 66;
    const radius = 56;
    const px = cx + radius * Math.cos(rad);
    const py = cy - radius * Math.sin(rad);
    alignPreviewDialPointer.style.left = `${px}px`;
    alignPreviewDialPointer.style.top = `${py}px`;
    alignPreviewDialText.style.transform = `translate(0, -50%) rotate(${-rotation}deg)`;
  };

  const syncJustifyLastLineAvailability = (): void => {
    projectDisabledState(justifyLastLineCk.input, draft.align !== 'distributed', null);
    justifyLastLineCk.wrap.classList.toggle(
      'fc-fmtdlg__check--muted',
      justifyLastLineCk.input.disabled,
    );
  };

  const sync = (): void => {
    const hKey: 'default' | CellAlign = draft.align ?? 'default';
    for (const [id, r] of hAlignRadios) r.checked = id === hKey;
    hAlignSelect.value = hKey;
    const vKey: 'default' | CellVAlign = draft.vAlign ?? 'default';
    for (const [id, r] of vAlignRadios) r.checked = id === vKey;
    vAlignSelect.value = vKey;
    wrapCk.input.checked = draft.wrap;
    justifyLastLineCk.input.checked = draft.justifyLastLine;
    syncJustifyLastLineAvailability();
    shrinkCk.input.checked = draft.shrinkToFit;
    indentInput.value = String(draft.indent);
    textDirectionSelect.value = draft.textDirection;
    rotationInput.value = String(draft.rotation);
    syncRotationDial();
  };

  const syncMixed = (): void => {
    markMixedCheck(wrapCk.input, isMixed('wrap'), 'wrap');
    markMixedCheck(justifyLastLineCk.input, isMixed('justifyLastLine'), 'justifyLastLine');
    markMixedCheck(shrinkCk.input, isMixed('shrinkToFit'), 'shrinkToFit');
    if (isMixed('align')) {
      for (const radio of hAlignRadios.values()) radio.checked = false;
      hAlignSelect.value = '';
      markMixed(hAlignSelect, true, 'align');
    }
    if (isMixed('vAlign')) {
      for (const radio of vAlignRadios.values()) radio.checked = false;
      vAlignSelect.value = '';
      markMixed(vAlignSelect, true, 'vAlign');
    }
    for (const [field, input] of [
      ['indent', indentInput],
      ['rotation', rotationInput],
      ['textDirection', textDirectionSelect],
    ] as const) {
      if (isMixed(field)) {
        input.value = '';
        markMixed(input, true, field);
      }
    }
  };

  const hydrateMerge = (selected: MergeSelectionState, disabled: boolean): void => {
    mergeCk.input.checked = selected === 'merged';
    mergeCk.input.indeterminate = selected === 'mixed';
    merge.touched = false;
    merge.initialChecked = mergeCk.input.checked;
    merge.initialIndeterminate = mergeCk.input.indeterminate;
    const reason = disabled ? t.mergeCellsRequiresMultiCell : null;
    projectDisabledState(mergeCk.input, disabled, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: t.mergeCells,
    });
    projectDisabledReason(mergeCk.wrap, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: t.mergeCells,
    });
    mergeCk.wrap.classList.toggle('fc-fmtdlg__check--muted', mergeCk.input.disabled);
  };

  const mergeAction = (): 'merge' | 'unmerge' | null => {
    const changed =
      !mergeCk.input.disabled &&
      (mergeCk.input.checked !== merge.initialChecked ||
        (merge.touched && mergeCk.input.indeterminate !== merge.initialIndeterminate));
    if (!changed) return null;
    return mergeCk.input.checked ? 'merge' : 'unmerge';
  };

  // ── Events ─────────────────────────────────────────────────────────────
  const onHAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    touch('align');
    if (draft.align === 'distributed' && r.value !== 'distributed') touch('justifyLastLine');
    draft.align = r.value === 'default' ? undefined : (r.value as CellAlign);
    if (draft.align !== 'distributed') draft.justifyLastLine = false;
    syncJustifyLastLineAvailability();
    hAlignSelect.value = r.value;
    renderPreview();
  };
  const onVAlignChange = (e: Event): void => {
    const r = e.target as HTMLInputElement;
    if (!r.checked) return;
    touch('vAlign');
    draft.vAlign = r.value === 'default' ? undefined : (r.value as CellVAlign);
    vAlignSelect.value = r.value;
    renderPreview();
  };
  const onHAlignSelectChange = (): void => {
    const value = hAlignSelect.value as 'default' | CellAlign;
    touch('align');
    if (draft.align === 'distributed' && value !== 'distributed') touch('justifyLastLine');
    draft.align = value === 'default' ? undefined : value;
    if (draft.align !== 'distributed') draft.justifyLastLine = false;
    syncJustifyLastLineAvailability();
    for (const [id, r] of hAlignRadios) r.checked = id === value;
    renderPreview();
  };
  const onVAlignSelectChange = (): void => {
    const value = vAlignSelect.value as 'default' | CellVAlign;
    touch('vAlign');
    draft.vAlign = value === 'default' ? undefined : value;
    for (const [id, r] of vAlignRadios) r.checked = id === value;
    renderPreview();
  };
  const onWrapChange = (): void => {
    touch('wrap');
    draft.wrap = wrapCk.input.checked;
    renderPreview();
  };
  const onJustifyLastLineChange = (): void => {
    touch('justifyLastLine');
    draft.justifyLastLine = justifyLastLineCk.input.checked;
    renderPreview();
  };
  const onShrinkToFitChange = (): void => {
    touch('shrinkToFit');
    draft.shrinkToFit = shrinkCk.input.checked;
    renderPreview();
  };
  const onIndentInput = (): void => {
    touch('indent');
    const n = Number.parseInt(indentInput.value, 10);
    if (Number.isFinite(n)) draft.indent = Math.max(0, Math.min(15, n));
    renderPreview();
  };
  const onTextDirectionChange = (): void => {
    touch('textDirection');
    draft.textDirection = textDirectionSelect.value as TextDirection;
    renderPreview();
  };
  const onRotationInput = (): void => {
    touch('rotation');
    const n = Number.parseInt(rotationInput.value, 10);
    if (Number.isFinite(n)) draft.rotation = Math.max(-90, Math.min(90, n));
    renderPreview();
  };
  const onMergeChange = (): void => {
    merge.touched = true;
    mergeCk.input.indeterminate = false;
  };
  const onDialClick = (event: Event): void => {
    const target = event.target as Element | null;
    const dot = target?.closest<HTMLButtonElement>('[data-fc-angle]');
    if (!dot) return;
    const angle = Number.parseInt(dot.dataset.fcAngle ?? '0', 10);
    if (!Number.isFinite(angle)) return;
    touch('rotation');
    draft.rotation = Math.max(-90, Math.min(90, angle));
    rotationInput.value = String(draft.rotation);
    renderPreview();
  };

  for (const r of hAlignRadios.values()) on(r, 'change', onHAlignChange);
  for (const r of vAlignRadios.values()) on(r, 'change', onVAlignChange);
  on(hAlignSelect, 'change', onHAlignSelectChange);
  on(vAlignSelect, 'change', onVAlignSelectChange);
  on(wrapCk.input, 'change', onWrapChange);
  on(justifyLastLineCk.input, 'change', onJustifyLastLineChange);
  on(shrinkCk.input, 'change', onShrinkToFitChange);
  on(mergeCk.input, 'change', onMergeChange);
  on(indentInput, 'input', onIndentInput);
  on(textDirectionSelect, 'change', onTextDirectionChange);
  on(rotationInput, 'input', onRotationInput);
  on(alignPreviewDial, 'click', onDialClick);

  return { sync, syncMixed, syncRotationDial, hydrateMerge, mergeAction };
}
