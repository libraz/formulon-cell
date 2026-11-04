// Header / footer tab: built-in presets, custom left/center/right slots, and
// the odd/even, first page, scaling and margin-alignment options.
import type { PageSetup } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { appendDialogButton } from '../dialog-shell.js';
import { makeCheck, makeTextInput } from '../page-setup-dialog-dom.js';
import type { PageSetupTabContext } from './context.js';

type HeaderFooterPreset = 'none' | 'page' | 'sheet' | 'path' | 'custom';

export interface HeaderFooterTab {
  hydrate(setup: PageSetup): void;
  collect(): Partial<PageSetup>;
}

export function createHeaderFooterTab(
  panel: HTMLDivElement,
  ctx: PageSetupTabContext,
): HeaderFooterTab {
  const { t, mac, on } = ctx;

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
  // First-page and even-page sections share the same three-slot editor and
  // appear only while their option is checked.
  const firstHeaderTriple = makeTriple(t.firstPageHeader);
  const firstFooterTriple = makeTriple(t.firstPageFooter);
  const evenHeaderTriple = makeTriple(t.evenPageHeader);
  const evenFooterTriple = makeTriple(t.evenPageFooter);

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

    on(customButton, 'click', () => {
      select.value = 'custom';
      triple.leftInput.focus();
    });

    return { row, select, customButton };
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

  on(headerPreset.select, 'change', () => {
    applyPreset(headerPreset.select, headerTriple, {
      none: '',
      page: t.headerPageNumber,
      sheet: t.headerSheetName,
    });
  });
  on(footerPreset.select, 'change', () => {
    applyPreset(footerPreset.select, footerTriple, {
      none: '',
      page: t.footerPageNumber,
      path: t.footerWorkbookPath,
    });
  });

  // The Mac dialog edits the three slots only for a custom header or footer.
  const syncSlotVisibility = (): void => {
    if (!mac) return;
    headerTriple.legendRow.hidden = headerPreset.select.value !== 'custom';
    footerTriple.legendRow.hidden = footerPreset.select.value !== 'custom';
  };
  on(headerPreset.select, 'change', syncSlotVisibility);
  on(footerPreset.select, 'change', syncSlotVisibility);
  on(headerPreset.customButton, 'click', syncSlotVisibility);
  on(footerPreset.customButton, 'click', syncSlotVisibility);

  const markCustomOnEdit = (select: HTMLSelectElement): void => {
    select.value = 'custom';
  };
  for (const input of [headerTriple.leftInput, headerTriple.centerInput, headerTriple.rightInput]) {
    on(input, 'input', () => markCustomOnEdit(headerPreset.select));
  }
  for (const input of [footerTriple.leftInput, footerTriple.centerInput, footerTriple.rightInput]) {
    on(input, 'input', () => markCustomOnEdit(footerPreset.select));
  }

  const headerFooterOptionsRow = document.createElement('div');
  headerFooterOptionsRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const headerFooterOptionsTitle = document.createElement('span');
  headerFooterOptionsTitle.textContent = t.tabHeaderFooter;
  const headerFooterOptionsValue = document.createElement('span');
  headerFooterOptionsValue.className = 'fc-pgsetup__value fc-pgsetup__checks';
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

  const syncVariantRows = (): void => {
    firstHeaderTriple.legendRow.hidden = !differentFirstPage.input.checked;
    firstFooterTriple.legendRow.hidden = !differentFirstPage.input.checked;
    evenHeaderTriple.legendRow.hidden = !differentOddEven.input.checked;
    evenFooterTriple.legendRow.hidden = !differentOddEven.input.checked;
  };
  on(differentFirstPage.input, 'change', syncVariantRows);
  on(differentOddEven.input, 'change', syncVariantRows);
  syncVariantRows();

  // The Mac tab stacks each pop-up over a preview of the three slots, puts
  // both custom buttons on one row between the halves and lists the options
  // in a single untitled column.
  const previews: { el: HTMLDivElement; triple: typeof headerTriple }[] = [];
  const makePreview = (triple: typeof headerTriple): HTMLDivElement => {
    const el = document.createElement('div');
    el.className = 'fc-pgsetup__hf-preview';
    el.setAttribute('aria-hidden', 'true');
    el.append(
      document.createElement('span'),
      document.createElement('span'),
      document.createElement('span'),
    );
    previews.push({ el, triple });
    return el;
  };
  const syncPreviews = (): void => {
    for (const { el, triple } of previews) {
      const [left, center, right] = Array.from(el.children);
      if (left) left.textContent = triple.leftInput.value;
      if (center) center.textContent = triple.centerInput.value;
      if (right) right.textContent = triple.rightInput.value;
    }
  };
  const macParts: HTMLElement[] = [];
  if (mac) {
    panel.classList.add('fc-pgsetup__tab--mac');
    headerPreset.row.classList.add('fc-pgsetup__preset-row--mac');
    footerPreset.row.classList.add('fc-pgsetup__preset-row--mac');
    headerFooterOptionsRow.classList.add('fc-pgsetup__hf-options--mac');
    const buttons = document.createElement('div');
    buttons.className = 'fc-pgsetup__hf-buttons';
    buttons.append(headerPreset.customButton, footerPreset.customButton);
    macParts.push(
      headerPreset.row,
      makePreview(headerTriple),
      headerTriple.legendRow,
      buttons,
      footerPreset.row,
      makePreview(footerTriple),
      footerTriple.legendRow,
    );
    for (const triple of [headerTriple, footerTriple]) {
      for (const input of [triple.leftInput, triple.centerInput, triple.rightInput]) {
        on(input, 'input', syncPreviews);
      }
    }
    on(headerPreset.select, 'change', syncPreviews);
    on(footerPreset.select, 'change', syncPreviews);
  }

  panel.append(
    ...(mac
      ? macParts
      : [headerPreset.row, headerTriple.legendRow, footerPreset.row, footerTriple.legendRow]),
    headerFooterOptionsRow,
    firstHeaderTriple.legendRow,
    firstFooterTriple.legendRow,
    evenHeaderTriple.legendRow,
    evenFooterTriple.legendRow,
  );

  return {
    hydrate(setup) {
      headerTriple.leftInput.value = setup.headerLeft ?? '';
      headerTriple.centerInput.value = setup.headerCenter ?? '';
      headerTriple.rightInput.value = setup.headerRight ?? '';
      footerTriple.leftInput.value = setup.footerLeft ?? '';
      footerTriple.centerInput.value = setup.footerCenter ?? '';
      footerTriple.rightInput.value = setup.footerRight ?? '';
      firstHeaderTriple.leftInput.value = setup.firstHeaderLeft ?? '';
      firstHeaderTriple.centerInput.value = setup.firstHeaderCenter ?? '';
      firstHeaderTriple.rightInput.value = setup.firstHeaderRight ?? '';
      firstFooterTriple.leftInput.value = setup.firstFooterLeft ?? '';
      firstFooterTriple.centerInput.value = setup.firstFooterCenter ?? '';
      firstFooterTriple.rightInput.value = setup.firstFooterRight ?? '';
      evenHeaderTriple.leftInput.value = setup.evenHeaderLeft ?? '';
      evenHeaderTriple.centerInput.value = setup.evenHeaderCenter ?? '';
      evenHeaderTriple.rightInput.value = setup.evenHeaderRight ?? '';
      evenFooterTriple.leftInput.value = setup.evenFooterLeft ?? '';
      evenFooterTriple.centerInput.value = setup.evenFooterCenter ?? '';
      evenFooterTriple.rightInput.value = setup.evenFooterRight ?? '';
      differentOddEven.input.checked = setup.differentOddEvenPages === true;
      differentFirstPage.input.checked = setup.differentFirstPage === true;
      syncVariantRows();
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
      syncSlotVisibility();
      syncPreviews();
    },
    collect() {
      return {
        headerLeft: headerTriple.leftInput.value,
        headerCenter: headerTriple.centerInput.value,
        headerRight: headerTriple.rightInput.value,
        footerLeft: footerTriple.leftInput.value,
        footerCenter: footerTriple.centerInput.value,
        footerRight: footerTriple.rightInput.value,
        firstHeaderLeft: firstHeaderTriple.leftInput.value,
        firstHeaderCenter: firstHeaderTriple.centerInput.value,
        firstHeaderRight: firstHeaderTriple.rightInput.value,
        firstFooterLeft: firstFooterTriple.leftInput.value,
        firstFooterCenter: firstFooterTriple.centerInput.value,
        firstFooterRight: firstFooterTriple.rightInput.value,
        evenHeaderLeft: evenHeaderTriple.leftInput.value,
        evenHeaderCenter: evenHeaderTriple.centerInput.value,
        evenHeaderRight: evenHeaderTriple.rightInput.value,
        evenFooterLeft: evenFooterTriple.leftInput.value,
        evenFooterCenter: evenFooterTriple.centerInput.value,
        evenFooterRight: evenFooterTriple.rightInput.value,
        differentOddEvenPages: differentOddEven.input.checked,
        differentFirstPage: differentFirstPage.input.checked,
        scaleHeaderFooterWithDocument: scaleWithDocument.input.checked,
        alignHeaderFooterWithMargins: alignWithMargins.input.checked,
      };
    },
  };
}
