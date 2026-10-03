// Margins tab: page margins, printer non-printable insets, and page centering.
import type { printableMarginAdjustments } from '../../commands/print.js';
import type { PageSetup } from '../../store/store.js';
import { labelize, makeNumberInput, makeRow } from '../page-setup-dialog-dom.js';
import type { PageSetupTabContext } from './context.js';

export interface MarginsTab {
  /** Numeric inputs whose edits can change the printable-margin warning. */
  marginInputs: HTMLInputElement[];
  renderWarning(adjustments: ReturnType<typeof printableMarginAdjustments>): void;
  hydrate(setup: PageSetup): void;
  collect(): Pick<PageSetup, 'margins'> & Partial<PageSetup>;
}

export function createMarginsTab(panel: HTMLDivElement, ctx: PageSetupTabContext): MarginsTab {
  const { t } = ctx;

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
  marginGroup.append(
    labelize(t.marginTop, topInput),
    labelize(t.marginRight, rightInput),
    labelize(t.marginBottom, bottomInput),
    labelize(t.marginLeft, leftInput),
    labelize(t.marginHeader, headerMarginInput),
    labelize(t.marginFooter, footerMarginInput),
  );
  marginRow.valueCell.appendChild(marginGroup);
  panel.appendChild(marginRow.row);

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
  panel.appendChild(printableRow.row);
  const printableWarning = document.createElement('div');
  printableWarning.className = 'fc-pgsetup__warning';
  printableWarning.hidden = true;
  printableWarning.setAttribute('role', 'status');
  printableWarning.setAttribute('aria-live', 'polite');
  panel.appendChild(printableWarning);

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
  panel.appendChild(centerRow);

  return {
    marginInputs: [
      topInput,
      rightInput,
      bottomInput,
      leftInput,
      printableTopInput,
      printableRightInput,
      printableBottomInput,
      printableLeftInput,
    ],
    renderWarning(adjustments) {
      const marginLabels = {
        top: t.marginTop,
        right: t.marginRight,
        bottom: t.marginBottom,
        left: t.marginLeft,
      };
      printableWarning.hidden = adjustments.length === 0;
      printableWarning.textContent =
        adjustments.length === 0
          ? ''
          : `${t.printableMarginWarning} ${adjustments
              .map((item) => `${marginLabels[item.side]} ${item.effective}in`)
              .join(', ')}`;
    },
    hydrate(setup) {
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
    },
    collect() {
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
      return {
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
      };
    },
  };
}
