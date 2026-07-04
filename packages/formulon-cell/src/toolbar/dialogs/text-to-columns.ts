import { projectDisabledReason, projectDisabledState } from '../menu-a11y.js';
import {
  appendDialogActions,
  appendErrorRow,
  clearDialogError,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
  showDialogError,
} from './shell.js';

export interface TextToColumnsDialogStrings {
  title: string;
  dataType: string;
  delimited: string;
  fixedWidth: string;
  fixedWidthUnavailable: string;
  delimiters: string;
  tab: string;
  semicolon: string;
  comma: string;
  space: string;
  other: string;
  treatConsecutive: string;
  preview: string;
  noDelimited: string;
  ok: string;
  cancel: string;
}

export interface TextToColumnsDialogResult {
  delimiters: string[];
  collapseConsecutiveDelimiters: boolean;
}

export interface TextToColumnsDialogOptions {
  strings: TextToColumnsDialogStrings;
  initialDelimiters?: readonly string[];
  previewRows?: readonly string[];
}

const delimiterSpecs = (
  strings: TextToColumnsDialogStrings,
): readonly { value: string; label: string }[] => [
  { value: '\t', label: strings.tab },
  { value: ';', label: strings.semicolon },
  { value: ',', label: strings.comma },
  { value: ' ', label: strings.space },
];

const splitPreview = (
  rows: readonly string[],
  delimiters: readonly string[],
  collapse: boolean,
): string => {
  const active = delimiters.filter((delimiter) => delimiter !== '');
  if (active.length === 0) return rows.join('\n');
  const pattern =
    active.length === 1 && !collapse
      ? null
      : new RegExp(
          `(?:${active
            .map((delimiter) => delimiter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
            .join('|')})${collapse ? '+' : ''}`,
        );
  return rows
    .map((row) => {
      const parts = pattern ? row.split(pattern) : row.split(active[0] ?? '');
      return parts.join(' | ');
    })
    .join('\n');
};

export const showTextToColumnsDialog = (
  opts: TextToColumnsDialogOptions,
): Promise<TextToColumnsDialogResult | null> =>
  new Promise<TextToColumnsDialogResult | null>((resolve) => {
    const { strings } = opts;
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
    const initial = new Set(opts.initialDelimiters ?? [',']);

    const typeFieldset = document.createElement('fieldset');
    typeFieldset.className = 'fc-textcols__section fc-textcols__types';
    const typeLegend = document.createElement('legend');
    typeLegend.className = 'fc-tb__dlg__label';
    typeLegend.textContent = strings.dataType;
    const delimitedLabel = document.createElement('label');
    delimitedLabel.className = 'fc-fmtdlg__radio';
    const delimited = document.createElement('input');
    delimited.type = 'radio';
    delimited.name = 'fc-textcols-type';
    delimited.checked = true;
    delimitedLabel.append(delimited, document.createTextNode(strings.delimited));
    const fixedLabel = document.createElement('label');
    fixedLabel.className = 'fc-fmtdlg__radio';
    const fixed = document.createElement('input');
    fixed.type = 'radio';
    fixed.name = 'fc-textcols-type';
    projectDisabledState(fixed, true, strings.fixedWidthUnavailable, {
      ariaDescription: false,
      describedById: 'fc-textcols-fixed-width-unavailable',
    });
    projectDisabledReason(fixedLabel, strings.fixedWidthUnavailable, { ariaDescription: false });
    fixedLabel.append(fixed, document.createTextNode(strings.fixedWidth));
    const fixedNote = document.createElement('div');
    fixedNote.id = 'fc-textcols-fixed-width-unavailable';
    fixedNote.className = 'fc-tb__dlg__hint';
    fixedNote.textContent = strings.fixedWidthUnavailable;
    typeFieldset.append(typeLegend, delimitedLabel, fixedLabel, fixedNote);
    shell.body.appendChild(typeFieldset);

    const delimiterFieldset = document.createElement('fieldset');
    delimiterFieldset.className = 'fc-textcols__section fc-textcols__delimiters';
    const delimiterLegend = document.createElement('legend');
    delimiterLegend.className = 'fc-tb__dlg__label';
    delimiterLegend.textContent = strings.delimiters;
    delimiterFieldset.appendChild(delimiterLegend);
    const delimiterGrid = document.createElement('div');
    delimiterGrid.className = 'fc-textcols__delimiter-grid';
    delimiterGrid.setAttribute('role', 'group');
    delimiterGrid.setAttribute('aria-label', strings.delimiters);
    delimiterFieldset.appendChild(delimiterGrid);
    const checks: HTMLInputElement[] = [];
    for (const spec of delimiterSpecs(strings)) {
      const label = document.createElement('label');
      label.className = 'fc-textcols__delimiter';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = spec.value;
      checkbox.checked = initial.has(spec.value);
      checkbox.dataset.dialogField = `delimiter-${spec.value === '\t' ? 'tab' : spec.value}`;
      checks.push(checkbox);
      const text = document.createElement('span');
      text.textContent = spec.label;
      label.append(checkbox, text);
      delimiterGrid.appendChild(label);
    }
    const otherLabel = document.createElement('label');
    otherLabel.className = 'fc-textcols__other';
    const otherText = document.createElement('span');
    otherText.textContent = strings.other;
    const otherInput = document.createElement('input');
    otherInput.type = 'text';
    otherInput.className = 'fc-tb__dlg__input';
    otherInput.maxLength = 8;
    otherInput.dataset.dialogField = 'delimiter-other';
    const customInitial = [...initial].find(
      (delimiter) => !delimiterSpecs(strings).some((spec) => spec.value === delimiter),
    );
    if (customInitial) otherInput.value = customInitial;
    otherLabel.append(otherText, otherInput);
    delimiterGrid.appendChild(otherLabel);
    const collapseLabel = document.createElement('label');
    collapseLabel.className = 'fc-textcols__collapse';
    const collapse = document.createElement('input');
    collapse.type = 'checkbox';
    collapse.dataset.dialogField = 'collapse-consecutive';
    collapseLabel.append(collapse, document.createTextNode(strings.treatConsecutive));
    delimiterFieldset.appendChild(collapseLabel);
    shell.body.appendChild(delimiterFieldset);

    const previewWrap = document.createElement('section');
    previewWrap.className = 'fc-textcols__preview';
    const previewLabel = document.createElement('div');
    previewLabel.className = 'fc-textcols__preview-label fc-tb__dlg__label';
    previewLabel.textContent = strings.preview;
    const preview = document.createElement('pre');
    preview.className = 'fc-tb__dlg__preview';
    previewWrap.append(previewLabel, preview);
    shell.body.appendChild(previewWrap);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const selectedDelimiters = (): string[] => {
      const values = checks
        .filter((checkbox) => checkbox.checked)
        .map((checkbox) => checkbox.value);
      if (otherInput.value) values.push(otherInput.value);
      return values;
    };
    const updatePreview = (): void => {
      preview.textContent = splitPreview(
        opts.previewRows ?? [],
        selectedDelimiters(),
        collapse.checked,
      );
    };

    const lifecycle = installDialogLifecycle<TextToColumnsDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const delimiters = selectedDelimiters();
      if (delimiters.length === 0) {
        showDialogError(errorRow, strings.noDelimited);
        return;
      }
      lifecycle.finish({
        delimiters,
        collapseConsecutiveDelimiters: collapse.checked,
      });
    };

    for (const input of [...checks, collapse, otherInput]) {
      input.addEventListener('input', () => {
        clearDialogError(errorRow);
        updatePreview();
      });
      input.addEventListener('change', () => {
        clearDialogError(errorRow);
        updatePreview();
      });
    }
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    updatePreview();
    mountDialog(shell, checks.find((checkbox) => checkbox.checked) ?? checks[0] ?? okBtn);
  });
