// Page tab: printer profile, orientation, paper size, scaling, print quality
// and first page number.
import type { PageOrientation, PageSetup, PaperSize, PrintQuality } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { appendDialogButton } from '../dialog-shell.js';
import { makeNumberInput, makeRow, makeTextInput } from '../page-setup-dialog-dom.js';
import type { PageSetupTabContext } from './context.js';

const PAPER_SIZES: PaperSize[] = ['A4', 'A3', 'A5', 'letter', 'legal', 'tabloid'];
const ORIENTATIONS: PageOrientation[] = ['portrait', 'landscape'];

export interface PageTab {
  /** The printer row is shown or hidden by the dialog from the host's profiles. */
  printerRow: HTMLDivElement;
  printerSelect: HTMLSelectElement;
  printerRefreshBtn: HTMLButtonElement;
  printerStatus: HTMLSpanElement;
  hydrate(setup: PageSetup): void;
  collect(): Pick<PageSetup, 'orientation' | 'paperSize'> & Partial<PageSetup>;
}

export function createPageTab(panel: HTMLDivElement, ctx: PageSetupTabContext): PageTab {
  const { t } = ctx;

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
  const printerStatus = document.createElement('span');
  printerStatus.className = 'fc-pgsetup__status';
  printerStatus.setAttribute('role', 'status');
  printerStatus.setAttribute('aria-live', 'polite');
  printerRow.valueCell.insertBefore(printerSelect, printerRefreshBtn);
  printerRow.valueCell.appendChild(printerStatus);
  panel.appendChild(printerRow.row);

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
  panel.appendChild(orientRow.row);

  // ── Paper size ──────────────────────────────────────────────────────────
  const paperRow = makeRow(t.paperSize);
  const paperSelect = createDialogSelect(
    PAPER_SIZES.map((p) => ({ value: p, label: p })),
    'A4',
    { className: 'fc-pgsetup__select', ariaLabel: t.paperSize },
  );
  paperRow.valueCell.appendChild(paperSelect);
  panel.appendChild(paperRow.row);

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
  panel.appendChild(scalingRow);

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
  panel.appendChild(printQualityRow.row);

  const firstPageRow = makeRow(t.firstPageNumber);
  const firstPageInput = makeTextInput('', t.firstPageNumberPlaceholder);
  firstPageInput.setAttribute('aria-label', t.firstPageNumber);
  firstPageRow.valueCell.appendChild(firstPageInput);
  panel.appendChild(firstPageRow.row);

  return {
    printerRow: printerRow.row,
    printerSelect,
    printerRefreshBtn,
    printerStatus,
    hydrate(setup) {
      orientSelect.value = setup.orientation;
      paperSelect.value = setup.paperSize;
      const hasFit = (setup.fitWidth ?? 0) > 0 || (setup.fitHeight ?? 0) > 0;
      adjustInput.checked = !hasFit;
      fitInput.checked = hasFit;
      scaleInput.value = String(Math.round((setup.scale ?? 1) * 100));
      fitWidthInput.value = String(setup.fitWidth ?? 1);
      fitHeightInput.value = String(setup.fitHeight ?? 1);
      printQualitySelect.value = setup.printQuality ?? 'automatic';
      firstPageInput.value =
        typeof setup.firstPageNumber === 'number' ? String(setup.firstPageNumber) : '';
    },
    collect() {
      const orientation = (orientSelect.value as PageOrientation) ?? 'portrait';
      const paperSize = (paperSelect.value as PaperSize) ?? 'A4';
      const scaleRaw = Number.parseFloat(scaleInput.value);
      const scale = Number.isFinite(scaleRaw) && scaleRaw > 0 ? scaleRaw / 100 : 1;
      const fitW = Number.parseInt(fitWidthInput.value, 10);
      const fitH = Number.parseInt(fitHeightInput.value, 10);
      const firstPage = Number.parseInt(firstPageInput.value, 10);
      return {
        orientation,
        paperSize,
        scale,
        fitWidth: fitInput.checked && Number.isFinite(fitW) && fitW > 0 ? fitW : 0,
        fitHeight: fitInput.checked && Number.isFinite(fitH) && fitH > 0 ? fitH : 0,
        printQuality: printQualitySelect.value as PrintQuality,
        firstPageNumber:
          Number.isFinite(firstPage) && firstPage > 0 && firstPageInput.value.trim()
            ? firstPage
            : undefined,
      };
    },
  };
}
