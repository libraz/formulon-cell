import { createInteractionChipButton } from '../../interact/chip-button.js';
import { projectDisabledReason, projectDisabledState } from '../menu-a11y.js';
import { createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogButton,
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

/** Wording of the three-step wizard. `{step}` in `title` stands for the step number. */
export interface TextToColumnsWizardStrings {
  title: string;
  step1Intro: string;
  step1Prompt: string;
  delimitedHint: string;
  fixedWidth: string;
  fixedHint: string;
  preview: string;
  step2Prompt: string;
  /** Step 2 prompt for fixed-width data; falls back to `step2Prompt`. */
  fixedPrompt?: string;
  /** Instruction lines shown under the fixed-width prompt, one per line. */
  fixedSteps?: readonly string[];
  delimiters: string;
  other: string;
  qualifier: string;
  /** Bare word for the no-qualifier option; rendered in braces like the desktop list. */
  qualifierNone: string;
  step3Prompt: string;
  columnFormat: string;
  general: string;
  text: string;
  date: string;
  skip: string;
  destination: string;
  details: string;
  back: string;
  next: string;
  finish: string;
}

export interface TextToColumnsDialogResult {
  delimiters: string[];
  collapseConsecutiveDelimiters: boolean;
  /** Set by the wizard: the quote that keeps delimiters inside a field, or null for none. */
  textQualifier?: string | null;
  /** Set by the wizard for fixed-width data: positions that start a new column. */
  fixedBreaks?: number[];
}

export interface TextToColumnsDialogOptions {
  strings: TextToColumnsDialogStrings;
  /** When given, the dialog runs as the desktop three-step wizard. */
  wizard?: TextToColumnsWizardStrings;
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

interface DelimiterControls {
  grid: HTMLDivElement;
  checks: HTMLInputElement[];
  otherInput: HTMLInputElement;
  collapseLabel: HTMLLabelElement;
  collapse: HTMLInputElement;
  selected: () => string[];
}

/** The delimiter checkbox grid with its Other field, and the consecutive-delimiters checkbox. */
const buildDelimiterControls = (
  strings: TextToColumnsDialogStrings,
  initial: ReadonlySet<string>,
  config: { gridLabel: string; otherLabel: string; otherMaxLength: number },
): DelimiterControls => {
  const specs = delimiterSpecs(strings);
  const grid = document.createElement('div');
  grid.className = 'fc-textcols__delimiter-grid';
  grid.setAttribute('role', 'group');
  grid.setAttribute('aria-label', config.gridLabel);
  const checks: HTMLInputElement[] = [];
  for (const spec of specs) {
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
    grid.appendChild(label);
  }
  const otherLabel = document.createElement('label');
  otherLabel.className = 'fc-textcols__other';
  const otherText = document.createElement('span');
  otherText.textContent = config.otherLabel;
  const otherInput = document.createElement('input');
  otherInput.type = 'text';
  otherInput.className = 'fc-tb__dlg__input';
  otherInput.maxLength = config.otherMaxLength;
  otherInput.dataset.dialogField = 'delimiter-other';
  const customInitial = [...initial].find(
    (delimiter) => !specs.some((spec) => spec.value === delimiter),
  );
  if (customInitial) otherInput.value = customInitial;
  otherLabel.append(otherText, otherInput);
  grid.appendChild(otherLabel);
  const collapseLabel = document.createElement('label');
  collapseLabel.className = 'fc-textcols__collapse';
  const collapse = document.createElement('input');
  collapse.type = 'checkbox';
  collapse.dataset.dialogField = 'collapse-consecutive';
  collapseLabel.append(collapse, document.createTextNode(strings.treatConsecutive));
  const selected = (): string[] => {
    const values = checks.filter((c) => c.checked).map((c) => c.value);
    if (otherInput.value) values.push(otherInput.value);
    return values;
  };
  return { grid, checks, otherInput, collapseLabel, collapse, selected };
};

/** The labelled source-data preview section; returns the `pre` that shows the text. */
const appendPreview = (body: HTMLElement, label: string): HTMLPreElement => {
  const previewWrap = document.createElement('section');
  previewWrap.className = 'fc-textcols__preview';
  const previewLabel = document.createElement('div');
  previewLabel.className = 'fc-textcols__preview-label fc-tb__dlg__label';
  previewLabel.textContent = label;
  const preview = document.createElement('pre');
  preview.className = 'fc-tb__dlg__preview';
  previewWrap.append(previewLabel, preview);
  body.appendChild(previewWrap);
  return preview;
};

const QUALIFIERS = ['"', "'"] as const;

const showTextToColumnsWizard = (
  opts: TextToColumnsDialogOptions & { wizard: TextToColumnsWizardStrings },
): Promise<TextToColumnsDialogResult | null> =>
  new Promise<TextToColumnsDialogResult | null>((resolve) => {
    const { strings, wizard } = opts;
    const stepTitle = (n: number): string => wizard.title.replace('{step}', String(n));
    const shell = createDialogShell({ title: stepTitle(1), bodyVariant: 'app' });
    shell.panel.classList.add('fc-textcols-wizard');
    const prompt = (text: string): HTMLParagraphElement => {
      const p = document.createElement('p');
      p.className = 'fc-textcols__prompt';
      p.textContent = text;
      return p;
    };
    const stepPanel = (n: number): HTMLDivElement => {
      const panel = document.createElement('div');
      panel.className = 'fc-textcols__step';
      panel.dataset.step = String(n);
      panel.hidden = n !== 1;
      shell.body.appendChild(panel);
      return panel;
    };
    const step1 = stepPanel(1);
    const step2 = stepPanel(2);
    const step3 = stepPanel(3);

    // Step 1: kind of source data.
    const typeRow = (input: HTMLInputElement, label: string, hint: string): HTMLLabelElement => {
      const row = document.createElement('label');
      row.className = 'fc-fmtdlg__radio fc-textcols__type-row';
      const hintEl = document.createElement('span');
      hintEl.className = 'fc-tb__dlg__hint';
      hintEl.textContent = hint;
      row.append(input, document.createTextNode(label), hintEl);
      return row;
    };
    const delimited = document.createElement('input');
    delimited.type = 'radio';
    delimited.name = 'fc-textcols-type';
    delimited.checked = true;
    const fixed = document.createElement('input');
    fixed.type = 'radio';
    fixed.name = 'fc-textcols-type';
    step1.append(
      prompt(wizard.step1Intro),
      prompt(wizard.step1Prompt),
      typeRow(delimited, strings.delimited, wizard.delimitedHint),
      typeRow(fixed, wizard.fixedWidth, wizard.fixedHint),
    );

    // Step 2: delimiters, consecutive runs and the quote character.
    const initial = new Set(opts.initialDelimiters ?? ['\t']);
    const previewRows = opts.previewRows ?? [];
    const {
      grid: delimiterBox,
      checks,
      otherInput,
      collapseLabel,
      collapse,
      selected: selectedDelimiters,
    } = buildDelimiterControls(strings, initial, {
      gridLabel: wizard.delimiters,
      otherLabel: wizard.other,
      otherMaxLength: 1,
    });
    const qualifierLabel = document.createElement('label');
    qualifierLabel.className = 'fc-textcols__qualifier';
    const qualifier = createDialogSelect(
      [
        ...QUALIFIERS.map((q) => ({ value: q, label: q })),
        { value: '', label: `{${wizard.qualifierNone}}` },
      ],
      '"',
      { className: '', ariaLabel: wizard.qualifier },
    );
    const qualifierText = document.createElement('span');
    qualifierText.textContent = wizard.qualifier;
    qualifierLabel.append(qualifierText, qualifier);
    const delimiterLegend = document.createElement('div');
    delimiterLegend.className = 'fc-tb__dlg__label';
    delimiterLegend.textContent = wizard.delimiters;
    const step2Prompt = prompt(wizard.step2Prompt);
    // Fixed-width data: click between characters to set or clear a column break.
    const fixedBreaks = new Set<number>();
    const fixedEditor = document.createElement('div');
    fixedEditor.className = 'fc-textcols__fixed';
    fixedEditor.hidden = true;
    // The ruler runs one tick per character, longer every fifth, numbered every tenth.
    const rulerLength = Math.min(
      120,
      Math.max(0, ...previewRows.slice(0, 8).map((row) => row.length)),
    );
    if (wizard.fixedSteps?.length) {
      const steps = document.createElement('div');
      steps.className = 'fc-textcols__fixed-steps';
      for (const text of wizard.fixedSteps) {
        const line = document.createElement('p');
        line.className = 'fc-textcols__fixed-step';
        line.textContent = text;
        steps.appendChild(line);
      }
      fixedEditor.appendChild(steps);
    }
    if (rulerLength > 0) {
      const ruler = document.createElement('div');
      ruler.className = 'fc-textcols__ruler';
      for (let pos = 0; pos < rulerLength; pos += 1) {
        const tick = document.createElement('span');
        tick.className = 'fc-textcols__ruler-tick';
        if ((pos + 1) % 5 === 0) tick.classList.add('fc-textcols__ruler-tick--long');
        if ((pos + 1) % 10 === 0) {
          const label = document.createElement('span');
          label.className = 'fc-textcols__ruler-label';
          label.textContent = String(pos + 1);
          tick.appendChild(label);
        }
        ruler.appendChild(tick);
      }
      fixedEditor.appendChild(ruler);
    }
    const fixedRows = previewRows.slice(0, 8).map((row) => {
      const line = document.createElement('div');
      line.className = 'fc-textcols__fixed-row';
      const chars = [...row.slice(0, 120)].map((char, pos) => {
        const cell = createInteractionChipButton({
          className: 'fc-textcols__fixed-char',
          label: char === ' ' ? '\u00a0' : char,
          dataset: { pos: String(pos) },
        });
        cell.addEventListener('click', () => {
          if (pos === 0) return;
          if (fixedBreaks.has(pos)) fixedBreaks.delete(pos);
          else fixedBreaks.add(pos);
          syncFixedBreaks();
          updatePreview();
        });
        line.appendChild(cell);
        return cell;
      });
      fixedEditor.appendChild(line);
      return chars;
    });
    const syncFixedBreaks = (): void => {
      for (const chars of fixedRows) {
        for (const [pos, cell] of chars.entries()) {
          cell.classList.toggle('fc-textcols__fixed-char--break', fixedBreaks.has(pos));
        }
      }
    };
    step2.append(
      step2Prompt,
      fixedEditor,
      delimiterLegend,
      delimiterBox,
      collapseLabel,
      qualifierLabel,
    );

    // Step 3: column data format and destination; only General applies to the result today.
    const formatLegend = document.createElement('div');
    formatLegend.className = 'fc-tb__dlg__label';
    formatLegend.textContent = wizard.columnFormat;
    step3.append(prompt(wizard.step3Prompt), formatLegend);
    const formatName = 'fc-textcols-format';
    for (const [label, checkedByDefault] of [
      [wizard.general, true],
      [wizard.text, false],
      [wizard.date, false],
      [wizard.skip, false],
    ] as const) {
      const row = document.createElement('label');
      row.className = 'fc-fmtdlg__radio';
      const radio = document.createElement('input');
      radio.type = 'radio';
      radio.name = formatName;
      radio.checked = checkedByDefault;
      if (!checkedByDefault) projectDisabledState(radio, true, strings.fixedWidthUnavailable);
      row.append(radio, document.createTextNode(label));
      step3.appendChild(row);
    }
    const destRow = document.createElement('label');
    destRow.className = 'fc-textcols__destination';
    const destText = document.createElement('span');
    destText.textContent = wizard.destination;
    const destInput = document.createElement('input');
    destInput.type = 'text';
    destInput.className = 'fc-tb__dlg__input';
    projectDisabledState(destInput, true, strings.fixedWidthUnavailable);
    destRow.append(destText, destInput);
    step3.appendChild(destRow);
    const detailsBtn = appendDialogButton(step3, { label: wizard.details });
    projectDisabledState(detailsBtn, true, strings.fixedWidthUnavailable);

    // Preview of the source data, shared by every step.
    const preview = appendPreview(shell.body, wizard.preview);

    const errorRow = appendErrorRow(shell.body);
    const cancelBtn = appendDialogButton(shell.footer, { label: strings.cancel });
    const backBtn = appendDialogButton(shell.footer, { label: wizard.back });
    const nextBtn = appendDialogButton(shell.footer, { label: wizard.next });
    const finishBtn = appendDialogButton(shell.footer, {
      label: wizard.finish,
      variant: 'primary',
    });

    const isFixed = (): boolean => fixed.checked;
    function updatePreview(): void {
      preview.textContent = isFixed()
        ? previewRows
            .map((row) => {
              const cuts = [...fixedBreaks].filter((c) => c < row.length).sort((a, b) => a - b);
              const parts: string[] = [];
              let from = 0;
              for (const cut of [...cuts, row.length]) {
                parts.push(row.slice(from, cut));
                from = cut;
              }
              return parts.join(' | ');
            })
            .join('\n')
        : splitPreview(previewRows, selectedDelimiters(), collapse.checked);
    }
    /** Step 2 shows the delimiter choices or the break editor, whichever the data type needs. */
    const syncKind = (): void => {
      const fixedMode = isFixed();
      fixedEditor.hidden = !fixedMode;
      delimiterBox.hidden = fixedMode;
      collapseLabel.hidden = fixedMode;
      qualifierLabel.hidden = fixedMode;
      delimiterLegend.hidden = fixedMode;
      step2Prompt.textContent = fixedMode
        ? (wizard.fixedPrompt ?? wizard.step2Prompt)
        : wizard.step2Prompt;
      updatePreview();
    };
    delimited.addEventListener('change', syncKind);
    fixed.addEventListener('change', syncKind);
    let current = 1;
    const showStep = (n: number): void => {
      current = n;
      [step1, step2, step3].forEach((panel, i) => {
        panel.hidden = i + 1 !== n;
      });
      shell.header.textContent = stepTitle(n);
      projectDisabledState(backBtn, n === 1, null);
      projectDisabledState(nextBtn, n === 3, null);
      clearDialogError(errorRow);
    };

    const lifecycle = installDialogLifecycle<TextToColumnsDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onFinish(),
    });
    const onFinish = (): void => {
      if (isFixed()) {
        if (fixedBreaks.size === 0) {
          showStep(2);
          showDialogError(errorRow, strings.noDelimited);
          return;
        }
        lifecycle.finish({
          delimiters: [],
          collapseConsecutiveDelimiters: false,
          textQualifier: null,
          fixedBreaks: [...fixedBreaks].sort((a, b) => a - b),
        });
        return;
      }
      const delimiters = selectedDelimiters();
      if (delimiters.length === 0) {
        showStep(2);
        showDialogError(errorRow, strings.noDelimited);
        return;
      }
      lifecycle.finish({
        delimiters,
        collapseConsecutiveDelimiters: collapse.checked,
        textQualifier: qualifier.value === '' ? null : qualifier.value,
      });
    };
    for (const input of [...checks, collapse, otherInput]) {
      const refresh = (): void => {
        clearDialogError(errorRow);
        updatePreview();
      };
      input.addEventListener('input', refresh);
      input.addEventListener('change', refresh);
    }
    backBtn.addEventListener('click', () => showStep(Math.max(1, current - 1)));
    nextBtn.addEventListener('click', () => showStep(Math.min(3, current + 1)));
    finishBtn.addEventListener('click', onFinish);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    showStep(1);
    updatePreview();
    mountDialog(shell, delimited);
  });

export const showTextToColumnsDialog = (
  opts: TextToColumnsDialogOptions,
): Promise<TextToColumnsDialogResult | null> => {
  const { wizard } = opts;
  if (wizard) return showTextToColumnsWizard({ ...opts, wizard });
  return showTextToColumnsSinglePage(opts);
};

const showTextToColumnsSinglePage = (
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
    const {
      grid: delimiterGrid,
      checks,
      otherInput,
      collapseLabel,
      collapse,
      selected: selectedDelimiters,
    } = buildDelimiterControls(strings, initial, {
      gridLabel: strings.delimiters,
      otherLabel: strings.other,
      otherMaxLength: 8,
    });
    delimiterFieldset.append(delimiterGrid, collapseLabel);
    shell.body.appendChild(delimiterFieldset);

    const preview = appendPreview(shell.body, strings.preview);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

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
      buttons: { okBtn, cancelBtn },
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

    updatePreview();
    mountDialog(shell, checks.find((checkbox) => checkbox.checked) ?? checks[0] ?? okBtn);
  });
