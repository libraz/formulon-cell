// Page Setup dialog. Lets the user edit the active sheet's `PageSetup` —
// orientation, paper size, margins (inches), header / footer slots, print
// titles, scale, gridlines / headings toggles. OK pushes the resulting patch
// through `mutators.setPageSetup` wrapped in a single history entry so undo
// reverts the whole apply atomically.
import { type History, recordPageSetupChange } from '../commands/history.js';
import {
  colLetter,
  parsePrintAreas,
  parsePrintTitleCols,
  parsePrintTitleRows,
  printableMarginAdjustments,
} from '../commands/print.js';
import {
  normalizePrintableBounds,
  normalizePrinterProfileId,
  normalizePrinterProfiles,
  type PrinterProfile,
} from '../commands/printer-profile.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  defaultPageSetup,
  getPageSetup,
  mutators,
  type PageMargins,
  type PageOrientation,
  type PageSetup,
  type PaperSize,
  type PrintCellErrorsMode,
  type PrintCommentsMode,
  type PrintPageOrder,
  type PrintQuality,
  type SpreadsheetStore,
} from '../store/store.js';
import { appendDialogSelectOptions, createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatA1Range } from '../wrappers/toolbar-a1.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogIconButton,
  appendDialogTabPair,
  createDialogShell,
} from './dialog-shell.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface PageSetupDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
  /** Shared history. When provided, OK pushes one page-setup snapshot entry. */
  history?: History | null;
  /** Optional host hook for resolving printer non-printable bounds after the
   *  user changes paper size or orientation. Browser APIs do not expose these
   *  values, so Electron/native hosts can refresh their profile here. */
  resolvePrintableBounds?: (
    setup: PageSetup,
    sheet: number,
    previous: PageSetup,
    printerProfileId: string | undefined,
  ) => Partial<PageMargins> | null | undefined;
  getPrinterProfiles?: () => readonly PrinterProfile[] | undefined;
  getPrinterProfileId?: () => string | undefined;
  setPrinterProfileId?: (next: string | undefined) => void;
  refreshPrinterProfiles?: () =>
    | readonly PrinterProfile[]
    | undefined
    | Promise<readonly PrinterProfile[] | undefined>;
}

export interface PageSetupDialogHandle {
  open(tab?: PageSetupDialogTab): void;
  close(): void;
  detach(): void;
}

const PAPER_SIZES: PaperSize[] = ['A4', 'A3', 'A5', 'letter', 'legal', 'tabloid'];
const ORIENTATIONS: PageOrientation[] = ['portrait', 'landscape'];
export type PageSetupDialogTab = 'page' | 'margins' | 'headerFooter' | 'sheet';
type HeaderFooterPreset = 'none' | 'page' | 'sheet' | 'path' | 'custom';

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
  let openingPrinterProfileId = normalizePrinterProfileId(deps.getPrinterProfileId?.());

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
  const headerTitle = document.createElement('span');
  headerTitle.textContent = t.title;
  // Distinct from the footer Cancel button — otherwise Playwright's strict
  // `getByRole('button', { name: 'Cancel' })` finds both and throws.
  // Heuristic: if t.cancel is non-ASCII (likely Japanese) use the Japanese
  // "Close" — otherwise English.
  const closeLabel = Array.from(t.cancel).some((ch) => ch.charCodeAt(0) > 0x7f)
    ? '閉じる'
    : 'Close';
  header.appendChild(headerTitle);
  const headerCloseBtn = appendDialogIconButton(header, {
    label: '',
    ariaLabel: closeLabel,
    baseClass: 'fc-fmtdlg__close',
  });
  panel.appendChild(header);

  const tabsStrip = document.createElement('div');
  tabsStrip.className = 'fc-fmtdlg__tabs';
  tabsStrip.setAttribute('role', 'tablist');
  tabsStrip.setAttribute('aria-label', t.title);
  panel.appendChild(tabsStrip);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  const tabDefs: { id: PageSetupDialogTab; label: string }[] = [
    { id: 'page', label: t.tabPage },
    { id: 'margins', label: t.tabMargins },
    { id: 'headerFooter', label: t.tabHeaderFooter },
    { id: 'sheet', label: t.tabSheet },
  ];
  const tabButtons = new Map<PageSetupDialogTab, HTMLButtonElement>();
  const tabPanels = new Map<PageSetupDialogTab, HTMLDivElement>();
  for (const def of tabDefs) {
    const { button, panel: tabPanel } = appendDialogTabPair(tabsStrip, body, {
      id: def.id,
      label: def.label,
      tabId: `fc-pgsetup-tab-${def.id}`,
      panelId: `fc-pgsetup-panel-${def.id}`,
      panelClass: 'fc-fmtdlg__panel-tab fc-pgsetup__tab-panel',
      tabDatasetKey: 'pgsetupTab',
      panelDatasetKey: 'pgsetupTab',
    });
    tabButtons.set(def.id, button);
    tabPanels.set(def.id, tabPanel);
  }

  const pagePanel = tabPanels.get('page') as HTMLDivElement;
  const marginsPanel = tabPanels.get('margins') as HTMLDivElement;
  const headerFooterPanel = tabPanels.get('headerFooter') as HTMLDivElement;
  const sheetPanel = tabPanels.get('sheet') as HTMLDivElement;

  const printerRow = makeRow(t.printerProfile);
  const printerSelect = createDialogSelect([], '', {
    className: 'fc-pgsetup__select',
    ariaLabel: t.printerProfile,
  });
  printerSelect.dataset.pgsetupPrinter = 'true';
  const printerRefreshBtn = appendDialogButton(printerRow.valueCell, {
    label: t.printerProfileRefresh,
    baseClass: 'fc-pgsetup__mini-btn',
  });
  printerRefreshBtn.hidden = !deps.refreshPrinterProfiles;
  const printerStatus = document.createElement('span');
  printerStatus.className = 'fc-pgsetup__status';
  printerStatus.setAttribute('role', 'status');
  printerStatus.setAttribute('aria-live', 'polite');
  printerRow.valueCell.insertBefore(printerSelect, printerRefreshBtn);
  printerRow.valueCell.appendChild(printerStatus);
  pagePanel.appendChild(printerRow.row);

  // ── Orientation ─────────────────────────────────────────────────────────
  const orientRow = makeRow(t.orientation);
  const orientSelect = createDialogSelect(
    ORIENTATIONS.map((o) => ({
      value: o,
      label: o === 'portrait' ? t.orientPortrait : t.orientLandscape,
    })),
    'portrait',
    { className: 'fc-pgsetup__select', ariaLabel: t.orientation },
  );
  orientRow.valueCell.appendChild(orientSelect);
  pagePanel.appendChild(orientRow.row);

  // ── Paper size ──────────────────────────────────────────────────────────
  const paperRow = makeRow(t.paperSize);
  const paperSelect = createDialogSelect(
    PAPER_SIZES.map((p) => ({ value: p, label: p })),
    'A4',
    { className: 'fc-pgsetup__select', ariaLabel: t.paperSize },
  );
  paperRow.valueCell.appendChild(paperSelect);
  pagePanel.appendChild(paperRow.row);

  // ── Scaling ─────────────────────────────────────────────────────────────
  const scalingRow = document.createElement('div');
  scalingRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const scalingTitle = document.createElement('span');
  scalingTitle.textContent = t.scaling;
  const scalingValue = document.createElement('span');
  scalingValue.className = 'fc-pgsetup__value fc-pgsetup__scaling';
  const adjustLabel = document.createElement('label');
  adjustLabel.className = 'fc-fmtdlg__check';
  const adjustInput = document.createElement('input');
  adjustInput.type = 'radio';
  adjustInput.name = 'fc-pgsetup-scaling';
  adjustInput.value = 'adjust';
  adjustInput.setAttribute('aria-label', t.adjustTo);
  const adjustText = document.createElement('span');
  adjustText.textContent = t.adjustTo;
  const scaleInput = makeNumberInput(100, 1, 10, 400);
  scaleInput.setAttribute('aria-label', t.scale);
  const percentText = document.createElement('span');
  percentText.textContent = t.percentNormalSize;
  adjustLabel.append(adjustInput, adjustText, scaleInput, percentText);

  const fitLabel = document.createElement('label');
  fitLabel.className = 'fc-fmtdlg__check';
  const fitInput = document.createElement('input');
  fitInput.type = 'radio';
  fitInput.name = 'fc-pgsetup-scaling';
  fitInput.value = 'fit';
  fitInput.setAttribute('aria-label', t.fitTo);
  const fitText = document.createElement('span');
  fitText.textContent = t.fitTo;
  const fitWidthInput = makeNumberInput(1, 1, 0, 99);
  fitWidthInput.setAttribute('aria-label', t.fitWidth);
  const pagesWideText = document.createElement('span');
  pagesWideText.textContent = t.pagesWideBy;
  const fitHeightInput = makeNumberInput(1, 1, 0, 99);
  fitHeightInput.setAttribute('aria-label', t.fitHeight);
  const tallText = document.createElement('span');
  tallText.textContent = t.tall;
  fitLabel.append(fitInput, fitText, fitWidthInput, pagesWideText, fitHeightInput, tallText);
  scalingValue.append(adjustLabel, fitLabel);
  scalingRow.append(scalingTitle, scalingValue);
  pagePanel.appendChild(scalingRow);

  const printQualityRow = makeRow(t.printQuality);
  const printQualitySelect = createDialogSelect(
    [
      { value: 'automatic', label: t.printQualityAutomatic },
      { value: '300', label: '300 dpi' },
      { value: '600', label: '600 dpi' },
      { value: '1200', label: '1200 dpi' },
    ],
    'automatic',
    { className: 'fc-pgsetup__select', ariaLabel: t.printQuality },
  );
  printQualityRow.valueCell.appendChild(printQualitySelect);
  pagePanel.appendChild(printQualityRow.row);

  const firstPageRow = makeRow(t.firstPageNumber);
  const firstPageInput = makeTextInput('', t.firstPageNumberPlaceholder);
  firstPageInput.setAttribute('aria-label', t.firstPageNumber);
  firstPageRow.valueCell.appendChild(firstPageInput);
  pagePanel.appendChild(firstPageRow.row);

  // ── Margins ─────────────────────────────────────────────────────────────
  const marginRow = makeRow(t.margins);
  const marginGroup = document.createElement('div');
  marginGroup.className = 'fc-pgsetup__margins';
  const topInput = makeNumberInput(0.7);
  topInput.setAttribute('aria-label', t.marginTop);
  const rightInput = makeNumberInput(0.7);
  rightInput.setAttribute('aria-label', t.marginRight);
  const bottomInput = makeNumberInput(0.7);
  bottomInput.setAttribute('aria-label', t.marginBottom);
  const leftInput = makeNumberInput(0.7);
  leftInput.setAttribute('aria-label', t.marginLeft);
  const headerMarginInput = makeNumberInput(0.3);
  headerMarginInput.setAttribute('aria-label', t.marginHeader);
  const footerMarginInput = makeNumberInput(0.3);
  footerMarginInput.setAttribute('aria-label', t.marginFooter);
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
    labelize(t.marginHeader, headerMarginInput),
    labelize(t.marginFooter, footerMarginInput),
  );
  marginRow.valueCell.appendChild(marginGroup);
  marginsPanel.appendChild(marginRow.row);

  const printableRow = makeRow(t.printerMargins);
  const printableGroup = document.createElement('div');
  printableGroup.className = 'fc-pgsetup__margins fc-pgsetup__printable';
  const printableTopInput = makeNumberInput(0);
  printableTopInput.setAttribute('aria-label', t.printableTop);
  const printableRightInput = makeNumberInput(0);
  printableRightInput.setAttribute('aria-label', t.printableRight);
  const printableBottomInput = makeNumberInput(0);
  printableBottomInput.setAttribute('aria-label', t.printableBottom);
  const printableLeftInput = makeNumberInput(0);
  printableLeftInput.setAttribute('aria-label', t.printableLeft);
  printableGroup.append(
    labelize(t.marginTop, printableTopInput),
    labelize(t.marginRight, printableRightInput),
    labelize(t.marginBottom, printableBottomInput),
    labelize(t.marginLeft, printableLeftInput),
  );
  printableRow.valueCell.appendChild(printableGroup);
  marginsPanel.appendChild(printableRow.row);
  const printableWarning = document.createElement('div');
  printableWarning.className = 'fc-pgsetup__warning';
  printableWarning.hidden = true;
  printableWarning.setAttribute('role', 'status');
  printableWarning.setAttribute('aria-live', 'polite');
  marginsPanel.appendChild(printableWarning);

  const centerRow = document.createElement('div');
  centerRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const centerTitle = document.createElement('span');
  centerTitle.textContent = t.centerOnPage;
  const centerValue = document.createElement('span');
  centerValue.className = 'fc-pgsetup__value fc-pgsetup__center';
  const centerHLabel = document.createElement('label');
  centerHLabel.className = 'fc-fmtdlg__check';
  const centerHInput = document.createElement('input');
  centerHInput.type = 'checkbox';
  centerHInput.setAttribute('aria-label', t.centerHorizontally);
  const centerHText = document.createElement('span');
  centerHText.textContent = t.centerHorizontally;
  centerHLabel.append(centerHInput, centerHText);
  const centerVLabel = document.createElement('label');
  centerVLabel.className = 'fc-fmtdlg__check';
  const centerVInput = document.createElement('input');
  centerVInput.type = 'checkbox';
  centerVInput.setAttribute('aria-label', t.centerVertically);
  const centerVText = document.createElement('span');
  centerVText.textContent = t.centerVertically;
  centerVLabel.append(centerVInput, centerVText);
  centerValue.append(centerHLabel, centerVLabel);
  centerRow.append(centerTitle, centerValue);
  marginsPanel.appendChild(centerRow);

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
    lInp.setAttribute('aria-label', `${legendText} ${t.slotLeftPlaceholder}`);
    const cInp = makeTextInput('', t.slotCenterPlaceholder);
    cInp.setAttribute('aria-label', `${legendText} ${t.slotCenterPlaceholder}`);
    const rInp = makeTextInput('', t.slotRightPlaceholder);
    rInp.setAttribute('aria-label', `${legendText} ${t.slotRightPlaceholder}`);
    wrap.append(lInp, cInp, rInp);
    legendRow.append(lbl, wrap);
    return { legendRow, leftInput: lInp, centerInput: cInp, rightInput: rInp };
  };
  const headerTriple = makeTriple(t.headerLabel);
  const footerTriple = makeTriple(t.footerLabel);

  const makeHeaderFooterPresetRow = (
    label: string,
    customLabel: string,
    options: { value: HeaderFooterPreset; label: string }[],
    triple: typeof headerTriple,
  ): { row: HTMLDivElement; select: HTMLSelectElement; customButton: HTMLButtonElement } => {
    const row = document.createElement('div');
    row.className = 'fc-pgsetup__row fc-fmtdlg__row fc-pgsetup__preset-row';
    const labelSpan = document.createElement('span');
    labelSpan.textContent = label;
    const valueCell = document.createElement('span');
    valueCell.className = 'fc-pgsetup__value';
    const select = createDialogSelect(options, options[0]?.value ?? '', {
      className: 'fc-pgsetup__select',
      ariaLabel: label,
    });
    const customButton = appendDialogButton(valueCell, {
      label: customLabel,
      baseClass: 'fc-fmtdlg__btn',
      secondaryClass: 'fc-pgsetup__custom-btn',
      variant: 'secondary',
    });
    customButton.setAttribute('aria-label', customLabel);
    valueCell.insertBefore(select, customButton);
    row.append(labelSpan, valueCell);

    customButton.addEventListener('click', () => {
      select.value = 'custom';
      triple.leftInput.focus();
    });

    return { row, select, customButton };
  };

  const makeOptionSelect = (
    label: string,
    options: { value: string; label: string }[],
  ): HTMLSelectElement => {
    return createDialogSelect(options, options[0]?.value ?? '', {
      className: 'fc-pgsetup__select',
      ariaLabel: label,
    });
  };

  const headerPreset = makeHeaderFooterPresetRow(
    t.headerBuiltin,
    t.customHeader,
    [
      { value: 'none', label: t.headerNone },
      { value: 'page', label: t.headerPageNumber },
      { value: 'sheet', label: t.headerSheetName },
      { value: 'custom', label: t.customHeader },
    ],
    headerTriple,
  );
  const footerPreset = makeHeaderFooterPresetRow(
    t.footerBuiltin,
    t.customFooter,
    [
      { value: 'none', label: t.headerNone },
      { value: 'page', label: t.footerPageNumber },
      { value: 'path', label: t.footerWorkbookPath },
      { value: 'custom', label: t.customFooter },
    ],
    footerTriple,
  );

  const applyPreset = (
    select: HTMLSelectElement,
    triple: typeof headerTriple,
    centerValueByPreset: Partial<Record<HeaderFooterPreset, string>>,
  ): void => {
    if (select.value === 'custom') {
      triple.leftInput.focus();
      return;
    }
    triple.leftInput.value = '';
    triple.centerInput.value = centerValueByPreset[select.value as HeaderFooterPreset] ?? '';
    triple.rightInput.value = '';
  };

  headerPreset.select.addEventListener('change', () => {
    applyPreset(headerPreset.select, headerTriple, {
      none: '',
      page: t.headerPageNumber,
      sheet: t.headerSheetName,
    });
  });
  footerPreset.select.addEventListener('change', () => {
    applyPreset(footerPreset.select, footerTriple, {
      none: '',
      page: t.footerPageNumber,
      path: t.footerWorkbookPath,
    });
  });

  const markCustomOnEdit = (select: HTMLSelectElement): void => {
    select.value = 'custom';
  };
  for (const input of [headerTriple.leftInput, headerTriple.centerInput, headerTriple.rightInput]) {
    input.addEventListener('input', () => markCustomOnEdit(headerPreset.select));
  }
  for (const input of [footerTriple.leftInput, footerTriple.centerInput, footerTriple.rightInput]) {
    input.addEventListener('input', () => markCustomOnEdit(footerPreset.select));
  }

  const headerFooterOptionsRow = document.createElement('div');
  headerFooterOptionsRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const headerFooterOptionsTitle = document.createElement('span');
  headerFooterOptionsTitle.textContent = t.tabHeaderFooter;
  const headerFooterOptionsValue = document.createElement('span');
  headerFooterOptionsValue.className = 'fc-pgsetup__value fc-pgsetup__checks';
  const makeCheck = (label: string): { labelEl: HTMLLabelElement; input: HTMLInputElement } => {
    const labelEl = document.createElement('label');
    labelEl.className = 'fc-fmtdlg__check';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.setAttribute('aria-label', label);
    const text = document.createElement('span');
    text.textContent = label;
    labelEl.append(input, text);
    return { labelEl, input };
  };
  const differentOddEven = makeCheck(t.differentOddEvenPages);
  const differentFirstPage = makeCheck(t.differentFirstPage);
  const scaleWithDocument = makeCheck(t.scaleWithDocument);
  const alignWithMargins = makeCheck(t.alignWithPageMargins);
  headerFooterOptionsValue.append(
    differentOddEven.labelEl,
    differentFirstPage.labelEl,
    scaleWithDocument.labelEl,
    alignWithMargins.labelEl,
  );
  headerFooterOptionsRow.append(headerFooterOptionsTitle, headerFooterOptionsValue);

  headerFooterPanel.append(
    headerPreset.row,
    headerTriple.legendRow,
    footerPreset.row,
    footerTriple.legendRow,
    headerFooterOptionsRow,
  );

  // ── Sheet: print area / titles ──────────────────────────────────────────
  const printAreaRow = makeRow(t.printArea);
  const printAreaInput = makeTextInput('', t.printAreaPlaceholder);
  printAreaInput.setAttribute('aria-label', t.printArea);
  printAreaRow.valueCell.appendChild(printAreaInput);
  attachRangePickerButton(printAreaInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => formatA1Range(store.getState().selection.range),
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'page-setup-print-area',
  });
  sheetPanel.appendChild(printAreaRow.row);

  // ── Print titles ────────────────────────────────────────────────────────
  const titleRowsRow = makeRow(t.printTitleRows);
  const titleRowsInput = makeTextInput('', t.printTitleRowsPlaceholder);
  titleRowsInput.setAttribute('aria-label', t.printTitleRows);
  titleRowsRow.valueCell.appendChild(titleRowsInput);
  attachRangePickerButton(titleRowsInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => {
      const range = store.getState().selection.range;
      return `${range.r0 + 1}:${range.r1 + 1}`;
    },
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'page-setup-print-title-rows',
  });
  sheetPanel.appendChild(titleRowsRow.row);

  const titleColsRow = makeRow(t.printTitleCols);
  const titleColsInput = makeTextInput('', t.printTitleColsPlaceholder);
  titleColsInput.setAttribute('aria-label', t.printTitleCols);
  titleColsRow.valueCell.appendChild(titleColsInput);
  attachRangePickerButton(titleColsInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => {
      const range = store.getState().selection.range;
      return `${colLetter(range.c0)}:${colLetter(range.c1)}`;
    },
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'page-setup-print-title-cols',
  });
  sheetPanel.appendChild(titleColsRow.row);

  // ── Sheet: print options ────────────────────────────────────────────────
  const gridRow = document.createElement('div');
  gridRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const printOptionsTitle = document.createElement('span');
  printOptionsTitle.textContent = t.printOptions;
  const printOptionsValue = document.createElement('span');
  printOptionsValue.className = 'fc-pgsetup__value fc-pgsetup__checks';
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

  const blackWhiteLabel = document.createElement('label');
  blackWhiteLabel.className = 'fc-fmtdlg__check';
  const blackWhiteInput = document.createElement('input');
  blackWhiteInput.type = 'checkbox';
  blackWhiteInput.setAttribute('aria-label', t.blackAndWhite);
  const blackWhiteText = document.createElement('span');
  blackWhiteText.textContent = t.blackAndWhite;
  blackWhiteLabel.append(blackWhiteInput, blackWhiteText);

  const draftLabel = document.createElement('label');
  draftLabel.className = 'fc-fmtdlg__check';
  const draftInput = document.createElement('input');
  draftInput.type = 'checkbox';
  draftInput.setAttribute('aria-label', t.draftQuality);
  const draftText = document.createElement('span');
  draftText.textContent = t.draftQuality;
  draftLabel.append(draftInput, draftText);

  printOptionsValue.append(showGridLabel, blackWhiteLabel, draftLabel, showHeadLabel);
  gridRow.append(printOptionsTitle, printOptionsValue);
  sheetPanel.appendChild(gridRow);

  const commentsRow = makeRow(t.comments);
  const commentsSelect = makeOptionSelect(t.comments, [
    { value: 'none', label: t.commentsNone },
    { value: 'asDisplayed', label: t.commentsAsDisplayed },
    { value: 'endOfSheet', label: t.commentsEndOfSheet },
  ]);
  commentsRow.valueCell.appendChild(commentsSelect);
  sheetPanel.appendChild(commentsRow.row);

  const errorsRow = makeRow(t.cellErrorsAs);
  const errorsSelect = makeOptionSelect(t.cellErrorsAs, [
    { value: 'displayed', label: t.cellErrorsDisplayed },
    { value: 'blank', label: t.cellErrorsBlank },
    { value: 'dash', label: t.cellErrorsDash },
    { value: 'na', label: t.cellErrorsNA },
  ]);
  errorsRow.valueCell.appendChild(errorsSelect);
  sheetPanel.appendChild(errorsRow.row);

  const pageOrderRow = document.createElement('div');
  pageOrderRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const pageOrderTitle = document.createElement('span');
  pageOrderTitle.textContent = t.pageOrder;
  const pageOrderValue = document.createElement('span');
  pageOrderValue.className = 'fc-pgsetup__value';
  const downOverLabel = document.createElement('label');
  downOverLabel.className = 'fc-fmtdlg__check';
  const downOverInput = document.createElement('input');
  downOverInput.type = 'radio';
  downOverInput.name = 'fc-pgsetup-page-order';
  downOverInput.value = 'downThenOver';
  downOverInput.setAttribute('aria-label', t.pageOrderDownThenOver);
  const downOverText = document.createElement('span');
  downOverText.textContent = t.pageOrderDownThenOver;
  downOverLabel.append(downOverInput, downOverText);
  const overDownLabel = document.createElement('label');
  overDownLabel.className = 'fc-fmtdlg__check';
  const overDownInput = document.createElement('input');
  overDownInput.type = 'radio';
  overDownInput.name = 'fc-pgsetup-page-order';
  overDownInput.value = 'overThenDown';
  overDownInput.setAttribute('aria-label', t.pageOrderOverThenDown);
  const overDownText = document.createElement('span');
  overDownText.textContent = t.pageOrderOverThenDown;
  overDownLabel.append(overDownInput, overDownText);
  pageOrderValue.append(downOverLabel, overDownLabel);
  pageOrderRow.append(pageOrderTitle, pageOrderValue);
  sheetPanel.appendChild(pageOrderRow);

  const referenceError = document.createElement('div');
  referenceError.className = 'fc-pgsetup__error';
  referenceError.setAttribute('role', 'alert');
  referenceError.hidden = true;
  sheetPanel.appendChild(referenceError);

  // ── Footer / buttons ────────────────────────────────────────────────────
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
  });
  panel.appendChild(footer);

  /** Snapshot of the dialog values when it opened. Used by Cancel to revert
   *  inline edits and (more importantly) by OK to push a single history entry
   *  spanning the whole apply. */
  let opening: PageSetup = defaultPageSetup();
  let activeTab: PageSetupDialogTab = 'page';

  const referenceInputs = [printAreaInput, titleRowsInput, titleColsInput] as const;

  const clearReferenceError = (): void => {
    referenceError.hidden = true;
    referenceError.textContent = '';
    for (const input of referenceInputs) input.removeAttribute('aria-invalid');
  };

  const showReferenceError = (input: HTMLInputElement, message: string): void => {
    for (const candidate of referenceInputs) candidate.removeAttribute('aria-invalid');
    input.setAttribute('aria-invalid', 'true');
    referenceError.textContent = message;
    referenceError.hidden = false;
    setActiveTab('sheet');
    input.focus();
  };

  const validateReferenceInputs = (): boolean => {
    const area = printAreaInput.value.trim();
    if (area && !parsePrintAreas(area)) {
      showReferenceError(printAreaInput, t.invalidPrintArea);
      return false;
    }
    const titleRows = titleRowsInput.value.trim();
    if (titleRows && !parsePrintTitleRows(titleRows)) {
      showReferenceError(titleRowsInput, t.invalidPrintTitleRows);
      return false;
    }
    const titleCols = titleColsInput.value.trim();
    if (titleCols && !parsePrintTitleCols(titleCols)) {
      showReferenceError(titleColsInput, t.invalidPrintTitleCols);
      return false;
    }
    clearReferenceError();
    return true;
  };

  const setActiveTab = (id: PageSetupDialogTab): void => {
    activeTab = id;
    for (const [tabId, btn] of tabButtons) {
      btn.setAttribute('aria-selected', tabId === id ? 'true' : 'false');
      btn.tabIndex = tabId === id ? 0 : -1;
    }
    for (const [tabId, tabPanel] of tabPanels) {
      tabPanel.hidden = tabId !== id;
    }
  };

  const tabOrder = Array.from(tabButtons.keys());
  const focusTabByIndex = (index: number): void => {
    const next = tabOrder[(index + tabOrder.length) % tabOrder.length];
    if (!next) return;
    setActiveTab(next);
    tabButtons.get(next)?.focus();
  };

  const onTabClick = (event: MouseEvent): void => {
    const btn = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-pgsetup-tab]');
    const id = btn?.dataset.pgsetupTab as PageSetupDialogTab | undefined;
    if (!btn || !id) return;
    setActiveTab(id);
    btn.focus();
  };

  const onTabKeyDown = (event: KeyboardEvent): void => {
    const btn = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-pgsetup-tab]');
    const id = btn?.dataset.pgsetupTab as PageSetupDialogTab | undefined;
    const index = id ? tabOrder.indexOf(id) : -1;
    if (index < 0) return;
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      event.preventDefault();
      focusTabByIndex(index + 1);
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      event.preventDefault();
      focusTabByIndex(index - 1);
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusTabByIndex(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusTabByIndex(tabOrder.length - 1);
    }
  };

  const renderPrinterProfiles = (nextProfiles?: readonly PrinterProfile[]): void => {
    const profiles = deps.getPrinterProfiles?.() ?? [];
    const effectiveProfiles = normalizePrinterProfiles(nextProfiles ?? profiles) ?? [];
    const selected = normalizePrinterProfileId(deps.getPrinterProfileId?.()) ?? '';
    printerSelect.replaceChildren();
    appendDialogSelectOptions(printerSelect, [
      { value: '', label: t.printerProfileAutomatic },
      ...effectiveProfiles.flatMap((profile) =>
        profile.id ? [{ value: profile.id, label: profile.name || profile.id }] : [],
      ),
    ]);
    printerRow.row.hidden = effectiveProfiles.length === 0 && !deps.refreshPrinterProfiles;
    printerSelect.value = selected;
    if (printerSelect.value !== selected) printerSelect.value = '';
  };

  const refreshPrinterProfilesFromHost = async (): Promise<void> => {
    if (!deps.refreshPrinterProfiles) return;
    projectDisabledState(printerRefreshBtn, true, t.printerProfileRefreshInProgress, {
      datasetKey: 'disabledReason',
      titlePrefix: t.printerProfileRefresh,
    });
    printerStatus.textContent = '';
    try {
      const profiles = await deps.refreshPrinterProfiles();
      renderPrinterProfiles(profiles);
    } catch {
      printerStatus.textContent = t.printerProfileRefreshFailed;
    } finally {
      projectDisabledState(printerRefreshBtn, false, null, {
        datasetKey: 'disabledReason',
        titlePrefix: t.printerProfileRefresh,
      });
    }
  };

  const hydrateFrom = (setup: PageSetup): void => {
    openingPrinterProfileId = normalizePrinterProfileId(deps.getPrinterProfileId?.());
    renderPrinterProfiles();
    opening = { ...setup, margins: { ...setup.margins } };
    orientSelect.value = setup.orientation;
    paperSelect.value = setup.paperSize;
    topInput.value = String(setup.margins.top);
    rightInput.value = String(setup.margins.right);
    bottomInput.value = String(setup.margins.bottom);
    leftInput.value = String(setup.margins.left);
    printableTopInput.value = String(setup.printableBounds?.top ?? 0);
    printableRightInput.value = String(setup.printableBounds?.right ?? 0);
    printableBottomInput.value = String(setup.printableBounds?.bottom ?? 0);
    printableLeftInput.value = String(setup.printableBounds?.left ?? 0);
    headerMarginInput.value = String(setup.headerMargin ?? 0.3);
    footerMarginInput.value = String(setup.footerMargin ?? 0.3);
    centerHInput.checked = setup.centerHorizontally === true;
    centerVInput.checked = setup.centerVertically === true;
    headerTriple.leftInput.value = setup.headerLeft ?? '';
    headerTriple.centerInput.value = setup.headerCenter ?? '';
    headerTriple.rightInput.value = setup.headerRight ?? '';
    footerTriple.leftInput.value = setup.footerLeft ?? '';
    footerTriple.centerInput.value = setup.footerCenter ?? '';
    footerTriple.rightInput.value = setup.footerRight ?? '';
    differentOddEven.input.checked = setup.differentOddEvenPages === true;
    differentFirstPage.input.checked = setup.differentFirstPage === true;
    scaleWithDocument.input.checked = setup.scaleHeaderFooterWithDocument !== false;
    alignWithMargins.input.checked = setup.alignHeaderFooterWithMargins !== false;
    headerPreset.select.value =
      !headerTriple.leftInput.value &&
      !headerTriple.rightInput.value &&
      headerTriple.centerInput.value === ''
        ? 'none'
        : !headerTriple.leftInput.value &&
            !headerTriple.rightInput.value &&
            headerTriple.centerInput.value === t.headerPageNumber
          ? 'page'
          : !headerTriple.leftInput.value &&
              !headerTriple.rightInput.value &&
              headerTriple.centerInput.value === t.headerSheetName
            ? 'sheet'
            : 'custom';
    footerPreset.select.value =
      !footerTriple.leftInput.value &&
      !footerTriple.rightInput.value &&
      footerTriple.centerInput.value === ''
        ? 'none'
        : !footerTriple.leftInput.value &&
            !footerTriple.rightInput.value &&
            footerTriple.centerInput.value === t.footerPageNumber
          ? 'page'
          : !footerTriple.leftInput.value &&
              !footerTriple.rightInput.value &&
              footerTriple.centerInput.value === t.footerWorkbookPath
            ? 'path'
            : 'custom';
    printAreaInput.value = setup.printArea ?? '';
    titleRowsInput.value = setup.printTitleRows ?? '';
    titleColsInput.value = setup.printTitleCols ?? '';
    const hasFit = (setup.fitWidth ?? 0) > 0 || (setup.fitHeight ?? 0) > 0;
    adjustInput.checked = !hasFit;
    fitInput.checked = hasFit;
    scaleInput.value = String(Math.round((setup.scale ?? 1) * 100));
    fitWidthInput.value = String(setup.fitWidth ?? 1);
    fitHeightInput.value = String(setup.fitHeight ?? 1);
    printQualitySelect.value = setup.printQuality ?? 'automatic';
    firstPageInput.value =
      typeof setup.firstPageNumber === 'number' ? String(setup.firstPageNumber) : '';
    showGridInput.checked = setup.showGridlines === true;
    showHeadInput.checked = setup.showHeadings === true;
    blackWhiteInput.checked = setup.blackAndWhite === true;
    draftInput.checked = setup.draftQuality === true;
    commentsSelect.value = setup.comments ?? 'none';
    errorsSelect.value = setup.cellErrorsAs ?? 'displayed';
    downOverInput.checked = (setup.pageOrder ?? 'downThenOver') === 'downThenOver';
    overDownInput.checked = setup.pageOrder === 'overThenDown';
    clearReferenceError();
    setActiveTab('page');
  };

  const collectFromInputs = (): PageSetup => {
    const orientation = (orientSelect.value as PageOrientation) ?? 'portrait';
    const paperSize = (paperSelect.value as PaperSize) ?? 'A4';
    const top = Number.parseFloat(topInput.value) || 0;
    const right = Number.parseFloat(rightInput.value) || 0;
    const bottom = Number.parseFloat(bottomInput.value) || 0;
    const left = Number.parseFloat(leftInput.value) || 0;
    const headerMargin = Number.parseFloat(headerMarginInput.value);
    const footerMargin = Number.parseFloat(footerMarginInput.value);
    const printableTop = Math.max(0, Number.parseFloat(printableTopInput.value) || 0);
    const printableRight = Math.max(0, Number.parseFloat(printableRightInput.value) || 0);
    const printableBottom = Math.max(0, Number.parseFloat(printableBottomInput.value) || 0);
    const printableLeft = Math.max(0, Number.parseFloat(printableLeftInput.value) || 0);
    const hasPrintableBounds =
      printableTop > 0 || printableRight > 0 || printableBottom > 0 || printableLeft > 0;
    const scaleRaw = Number.parseFloat(scaleInput.value);
    const scale = Number.isFinite(scaleRaw) && scaleRaw > 0 ? scaleRaw / 100 : 1;
    const fitW = Number.parseInt(fitWidthInput.value, 10);
    const fitH = Number.parseInt(fitHeightInput.value, 10);
    const firstPage = Number.parseInt(firstPageInput.value, 10);
    const out: PageSetup = {
      orientation,
      paperSize,
      margins: { top, right, bottom, left },
      printableBounds: hasPrintableBounds
        ? {
            top: printableTop,
            right: printableRight,
            bottom: printableBottom,
            left: printableLeft,
          }
        : undefined,
      headerMargin: Number.isFinite(headerMargin) ? Math.max(0, headerMargin) : 0.3,
      footerMargin: Number.isFinite(footerMargin) ? Math.max(0, footerMargin) : 0.3,
      centerHorizontally: centerHInput.checked,
      centerVertically: centerVInput.checked,
      headerLeft: headerTriple.leftInput.value,
      headerCenter: headerTriple.centerInput.value,
      headerRight: headerTriple.rightInput.value,
      footerLeft: footerTriple.leftInput.value,
      footerCenter: footerTriple.centerInput.value,
      footerRight: footerTriple.rightInput.value,
      differentOddEvenPages: differentOddEven.input.checked,
      differentFirstPage: differentFirstPage.input.checked,
      scaleHeaderFooterWithDocument: scaleWithDocument.input.checked,
      alignHeaderFooterWithMargins: alignWithMargins.input.checked,
      printArea: printAreaInput.value.trim() || undefined,
      printTitleRows: titleRowsInput.value.trim() || undefined,
      printTitleCols: titleColsInput.value.trim() || undefined,
      scale,
      fitWidth: fitInput.checked && Number.isFinite(fitW) && fitW > 0 ? fitW : 0,
      fitHeight: fitInput.checked && Number.isFinite(fitH) && fitH > 0 ? fitH : 0,
      printQuality: printQualitySelect.value as PrintQuality,
      firstPageNumber:
        Number.isFinite(firstPage) && firstPage > 0 && firstPageInput.value.trim()
          ? firstPage
          : undefined,
      showGridlines: showGridInput.checked,
      showHeadings: showHeadInput.checked,
      blackAndWhite: blackWhiteInput.checked,
      draftQuality: draftInput.checked,
      comments: commentsSelect.value as PrintCommentsMode,
      cellErrorsAs: errorsSelect.value as PrintCellErrorsMode,
      pageOrder: (overDownInput.checked ? 'overThenDown' : 'downThenOver') as PrintPageOrder,
    };
    return out;
  };

  const updatePrintableWarning = (): void => {
    const marginLabels = {
      top: t.marginTop,
      right: t.marginRight,
      bottom: t.marginBottom,
      left: t.marginLeft,
    };
    const adjustments = printableMarginAdjustments(collectFromInputs());
    printableWarning.hidden = adjustments.length === 0;
    printableWarning.textContent =
      adjustments.length === 0
        ? ''
        : `${t.printableMarginWarning} ${adjustments
            .map((item) => `${marginLabels[item.side]} ${item.effective}in`)
            .join(', ')}`;
  };

  const onOk = (): void => {
    if (!validateReferenceInputs()) return;
    const sheet = store.getState().data.sheetIndex;
    const next = collectFromInputs();
    const nextPrinterProfileId = printerSelect.value || undefined;
    const paperChanged =
      next.paperSize !== opening.paperSize || next.orientation !== opening.orientation;
    const printerProfileChanged = nextPrinterProfileId !== openingPrinterProfileId;
    if (paperChanged || printerProfileChanged) {
      const resolved = deps.resolvePrintableBounds?.(next, sheet, opening, nextPrinterProfileId);
      if (resolved !== undefined) {
        next.printableBounds = resolved ? normalizePrintableBounds(resolved) : undefined;
      }
    }
    if (printerProfileChanged) deps.setPrinterProfileId?.(nextPrinterProfileId);
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
      if ((e.target as HTMLElement).tagName === 'BUTTON') return;
      // Enter inside a textbox triggers OK — spreadsheet parity.
      e.preventDefault();
      onOk();
    }
  };

  shell.on(tabsStrip, 'click', onTabClick as EventListener);
  shell.on(tabsStrip, 'keydown', onTabKeyDown as EventListener);
  shell.on(headerCloseBtn, 'click', onCancel);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(printerRefreshBtn, 'click', () => {
    void refreshPrinterProfilesFromHost();
  });
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);
  for (const input of referenceInputs) shell.on(input, 'input', clearReferenceError);
  for (const input of [
    topInput,
    rightInput,
    bottomInput,
    leftInput,
    printableTopInput,
    printableRightInput,
    printableBottomInput,
    printableLeftInput,
  ]) {
    shell.on(input, 'input', updatePrintableWarning);
  }

  const api: PageSetupDialogHandle = {
    open(tab: PageSetupDialogTab = 'page'): void {
      const sheet = store.getState().data.sheetIndex;
      hydrateFrom(getPageSetup(store.getState(), sheet));
      setActiveTab(tab);
      updatePrintableWarning();
      shell.open();
      requestAnimationFrame(() => {
        tabButtons.get(activeTab)?.focus();
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
