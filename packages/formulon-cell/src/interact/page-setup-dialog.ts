// Page Setup dialog. Lets the user edit the active sheet's `PageSetup` —
// orientation, paper size, margins (inches), header / footer slots, print
// titles, scale, gridlines / headings toggles. OK pushes the resulting patch
// through `mutators.setPageSetup` wrapped in a single history entry so undo
// reverts the whole apply atomically.
import { type History, recordPageSetupChange } from '../commands/history.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  defaultPageSetup,
  getPageSetup,
  mutators,
  type PageOrientation,
  type PageSetup,
  type PaperSize,
  type SpreadsheetStore,
} from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

export interface PageSetupDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Shared history. When provided, OK pushes one page-setup snapshot entry. */
  history?: History | null;
}

export interface PageSetupDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const PAPER_SIZES: PaperSize[] = ['A4', 'A3', 'A5', 'letter', 'legal', 'tabloid'];
const ORIENTATIONS: PageOrientation[] = ['portrait', 'landscape'];

function makeRow(label: string): { row: HTMLDivElement; valueCell: HTMLSpanElement } {
  const row = document.createElement('div');
  row.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const labelSpan = document.createElement('span');
  labelSpan.textContent = label;
  const valueCell = document.createElement('span');
  valueCell.className = 'fc-pgsetup__value';
  row.append(labelSpan, valueCell);
  return { row, valueCell };
}

function makeNumberInput(value: number, step = 0.1, min = 0, max = 99): HTMLInputElement {
  const inp = document.createElement('input');
  inp.type = 'number';
  inp.step = String(step);
  inp.min = String(min);
  inp.max = String(max);
  inp.value = String(value);
  inp.className = 'fc-pgsetup__num';
  return inp;
}

function makeTextInput(value: string, placeholder = ''): HTMLInputElement {
  const inp = document.createElement('input');
  inp.type = 'text';
  inp.value = value;
  inp.placeholder = placeholder;
  inp.autocomplete = 'off';
  inp.spellcheck = false;
  inp.className = 'fc-pgsetup__text';
  return inp;
}

export function attachPageSetupDialog(deps: PageSetupDialogDeps): PageSetupDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.pageSetup;

  const shell = createDialogShell({
    host,
    className: 'fc-pgsetup',
    ariaLabel: t.title,
    onDismiss: () => onCancel(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-pgsetup__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  // ── Orientation ─────────────────────────────────────────────────────────
  const orientRow = makeRow(t.orientation);
  const orientSelect = document.createElement('select');
  orientSelect.className = 'fc-pgsetup__select';
  for (const o of ORIENTATIONS) {
    const opt = document.createElement('option');
    opt.value = o;
    opt.textContent = o === 'portrait' ? t.orientPortrait : t.orientLandscape;
    orientSelect.appendChild(opt);
  }
  orientRow.valueCell.appendChild(orientSelect);
  body.appendChild(orientRow.row);

  // ── Paper size ──────────────────────────────────────────────────────────
  const paperRow = makeRow(t.paperSize);
  const paperSelect = document.createElement('select');
  paperSelect.className = 'fc-pgsetup__select';
  for (const p of PAPER_SIZES) {
    const opt = document.createElement('option');
    opt.value = p;
    opt.textContent = p;
    paperSelect.appendChild(opt);
  }
  paperRow.valueCell.appendChild(paperSelect);
  body.appendChild(paperRow.row);

  // ── Margins ─────────────────────────────────────────────────────────────
  const marginRow = makeRow(t.margins);
  const marginGroup = document.createElement('div');
  marginGroup.className = 'fc-pgsetup__margins';
  const topInput = makeNumberInput(0.7);
  const rightInput = makeNumberInput(0.7);
  const bottomInput = makeNumberInput(0.7);
  const leftInput = makeNumberInput(0.7);
  const labelize = (label: string, input: HTMLInputElement): HTMLLabelElement => {
    const lab = document.createElement('label');
    lab.className = 'fc-pgsetup__margin';
    const sp = document.createElement('span');
    sp.textContent = label;
    lab.append(sp, input);
    return lab;
  };
  marginGroup.append(
    labelize(t.marginTop, topInput),
    labelize(t.marginRight, rightInput),
    labelize(t.marginBottom, bottomInput),
    labelize(t.marginLeft, leftInput),
  );
  marginRow.valueCell.appendChild(marginGroup);
  body.appendChild(marginRow.row);

  // ── Header / footer ─────────────────────────────────────────────────────
  const makeTriple = (
    legendText: string,
  ): {
    legendRow: HTMLDivElement;
    leftInput: HTMLInputElement;
    centerInput: HTMLInputElement;
    rightInput: HTMLInputElement;
  } => {
    const legendRow = document.createElement('div');
    legendRow.className = 'fc-pgsetup__triple fc-fmtdlg__row';
    const lbl = document.createElement('span');
    lbl.textContent = legendText;
    const wrap = document.createElement('span');
    wrap.className = 'fc-pgsetup__value';
    const lInp = makeTextInput('', t.slotLeftPlaceholder);
    const cInp = makeTextInput('', t.slotCenterPlaceholder);
    const rInp = makeTextInput('', t.slotRightPlaceholder);
    wrap.append(lInp, cInp, rInp);
    legendRow.append(lbl, wrap);
    return { legendRow, leftInput: lInp, centerInput: cInp, rightInput: rInp };
  };
  const headerTriple = makeTriple(t.headerLabel);
  const footerTriple = makeTriple(t.footerLabel);
  body.append(headerTriple.legendRow, footerTriple.legendRow);

  // ── Print titles ────────────────────────────────────────────────────────
  const titleRowsRow = makeRow(t.printTitleRows);
  const titleRowsInput = makeTextInput('', t.printTitleRowsPlaceholder);
  titleRowsRow.valueCell.appendChild(titleRowsInput);
  body.appendChild(titleRowsRow.row);

  const titleColsRow = makeRow(t.printTitleCols);
  const titleColsInput = makeTextInput('', t.printTitleColsPlaceholder);
  titleColsRow.valueCell.appendChild(titleColsInput);
  body.appendChild(titleColsRow.row);

  // ── Scale + fit ─────────────────────────────────────────────────────────
  const scaleRow = makeRow(t.scale);
  const scaleInput = makeNumberInput(1, 0.05, 0.1, 4);
  scaleRow.valueCell.appendChild(scaleInput);
  body.appendChild(scaleRow.row);

  const fitWidthRow = makeRow(t.fitWidth);
  const fitWidthInput = makeNumberInput(0, 1, 0, 99);
  fitWidthRow.valueCell.appendChild(fitWidthInput);
  body.appendChild(fitWidthRow.row);

  const fitHeightRow = makeRow(t.fitHeight);
  const fitHeightInput = makeNumberInput(0, 1, 0, 99);
  fitHeightRow.valueCell.appendChild(fitHeightInput);
  body.appendChild(fitHeightRow.row);

  // ── Gridlines + headings ────────────────────────────────────────────────
  const gridRow = document.createElement('div');
  gridRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const showGridLabel = document.createElement('label');
  showGridLabel.className = 'fc-fmtdlg__check';
  const showGridInput = document.createElement('input');
  showGridInput.type = 'checkbox';
  const showGridText = document.createElement('span');
  showGridText.textContent = t.showGridlines;
  showGridLabel.append(showGridInput, showGridText);

  const showHeadLabel = document.createElement('label');
  showHeadLabel.className = 'fc-fmtdlg__check';
  const showHeadInput = document.createElement('input');
  showHeadInput.type = 'checkbox';
  const showHeadText = document.createElement('span');
  showHeadText.textContent = t.showHeadings;
  showHeadLabel.append(showHeadInput, showHeadText);

  gridRow.append(showGridLabel, showHeadLabel);
  body.appendChild(gridRow);

  // ── Footer / buttons ────────────────────────────────────────────────────
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-fmtdlg__btn';
  cancelBtn.textContent = t.cancel;
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  okBtn.textContent = t.ok;
  footer.append(cancelBtn, okBtn);
  panel.appendChild(footer);

  /** Snapshot of the dialog values when it opened. Used by Cancel to revert
   *  inline edits and (more importantly) by OK to push a single history entry
   *  spanning the whole apply. */
  let opening: PageSetup = defaultPageSetup();

  const hydrateFrom = (setup: PageSetup): void => {
    opening = { ...setup, margins: { ...setup.margins } };
    orientSelect.value = setup.orientation;
    paperSelect.value = setup.paperSize;
    topInput.value = String(setup.margins.top);
    rightInput.value = String(setup.margins.right);
    bottomInput.value = String(setup.margins.bottom);
    leftInput.value = String(setup.margins.left);
    headerTriple.leftInput.value = setup.headerLeft ?? '';
    headerTriple.centerInput.value = setup.headerCenter ?? '';
    headerTriple.rightInput.value = setup.headerRight ?? '';
    footerTriple.leftInput.value = setup.footerLeft ?? '';
    footerTriple.centerInput.value = setup.footerCenter ?? '';
    footerTriple.rightInput.value = setup.footerRight ?? '';
    titleRowsInput.value = setup.printTitleRows ?? '';
    titleColsInput.value = setup.printTitleCols ?? '';
    scaleInput.value = String(setup.scale ?? 1);
    fitWidthInput.value = String(setup.fitWidth ?? 0);
    fitHeightInput.value = String(setup.fitHeight ?? 0);
    showGridInput.checked = setup.showGridlines === true;
    showHeadInput.checked = setup.showHeadings === true;
  };

  const collectFromInputs = (): PageSetup => {
    const orientation = (orientSelect.value as PageOrientation) ?? 'portrait';
    const paperSize = (paperSelect.value as PaperSize) ?? 'A4';
    const top = Number.parseFloat(topInput.value) || 0;
    const right = Number.parseFloat(rightInput.value) || 0;
    const bottom = Number.parseFloat(bottomInput.value) || 0;
    const left = Number.parseFloat(leftInput.value) || 0;
    const scaleRaw = Number.parseFloat(scaleInput.value);
    const scale = Number.isFinite(scaleRaw) && scaleRaw > 0 ? scaleRaw : 1;
    const fitW = Number.parseInt(fitWidthInput.value, 10);
    const fitH = Number.parseInt(fitHeightInput.value, 10);
    const out: PageSetup = {
      orientation,
      paperSize,
      margins: { top, right, bottom, left },
      headerLeft: headerTriple.leftInput.value,
      headerCenter: headerTriple.centerInput.value,
      headerRight: headerTriple.rightInput.value,
      footerLeft: footerTriple.leftInput.value,
      footerCenter: footerTriple.centerInput.value,
      footerRight: footerTriple.rightInput.value,
      printTitleRows: titleRowsInput.value.trim() || undefined,
      printTitleCols: titleColsInput.value.trim() || undefined,
      scale,
      fitWidth: Number.isFinite(fitW) && fitW > 0 ? fitW : 0,
      fitHeight: Number.isFinite(fitH) && fitH > 0 ? fitH : 0,
      showGridlines: showGridInput.checked,
      showHeadings: showHeadInput.checked,
    };
    return out;
  };

  const onOk = (): void => {
    const sheet = store.getState().data.sheetIndex;
    const next = collectFromInputs();
    recordPageSetupChange(history, store, () => {
      mutators.setPageSetup(store, sheet, next);
    });
    api.close();
  };

  const onCancel = (): void => {
    // Cancel: rehydrate inputs so a follow-up open() doesn't show stale text,
    // and skip the slice mutation entirely. The opening snapshot is still in
    // the slice unchanged.
    hydrateFrom(opening);
    api.close();
  };

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
      return;
    }
    if (e.key === 'Enter') {
      // Enter inside a textbox triggers OK — spreadsheet parity.
      e.preventDefault();
      onOk();
    }
  };

  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const api: PageSetupDialogHandle = {
    open(): void {
      const sheet = store.getState().data.sheetIndex;
      hydrateFrom(getPageSetup(store.getState(), sheet));
      shell.open();
      requestAnimationFrame(() => {
        orientSelect.focus();
      });
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };

  return api;
}
