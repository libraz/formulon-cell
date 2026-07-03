// Fill Series subsystem: direction/mode types, the source-range collapse
// helper, the auto-direction heuristic, and the modal dialog that lets the
// user pick direction + series type. Kept self-contained so main.ts only
// owns the imperative `runFillSeries` / `applyFillSeries` glue.

import { dictionaries, type Strings } from '../../i18n/strings.js';
import type { Range } from '../../index.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from '../dialogs/shell.js';

export type RibbonFillDirection = 'down' | 'right' | 'up' | 'left';
export type RibbonFillSeriesMode = 'auto' | 'copy' | 'days' | 'weekdays' | 'months' | 'years';
export type FillSeriesDialogText = Strings['fillSeriesDialog'];

export const fillSeriesSourceRange = (range: Range, direction: RibbonFillDirection): Range => {
  if (direction === 'down') return { ...range, r1: range.r0 };
  if (direction === 'up') return { ...range, r0: range.r1 };
  if (direction === 'right') return { ...range, c1: range.c0 };
  return { ...range, c0: range.c1 };
};

export const inferFillSeriesDirection = (range: Range): RibbonFillDirection =>
  range.r1 > range.r0 ? 'down' : 'right';

export const makeFillSeriesRadio = <T extends string>(
  name: string,
  value: T,
  label: string,
  checked: boolean,
): HTMLLabelElement => {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__radio';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return wrap;
};

export const selectedFillSeriesRadio = <T extends string>(
  root: HTMLElement,
  name: string,
  fallback: T,
): T =>
  (root.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value as T | undefined) ??
  fallback;

export const showFillSeriesDialog = (
  range: Range,
  text: FillSeriesDialogText | 'ja' | 'en',
): Promise<{ direction: RibbonFillDirection; mode: RibbonFillSeriesMode } | null> => {
  return new Promise((resolve) => {
    const t = typeof text === 'string' ? dictionaries[text].fillSeriesDialog : text;
    const title = t.title;
    const shell = createDialogShell({ title });
    shell.panel.classList.add('fc-pastesp__panel');
    shell.body.classList.add('fc-pastesp__body');

    const cols = document.createElement('div');
    cols.className = 'fc-pastesp__cols';
    shell.body.appendChild(cols);

    const dirName = `app-fill-series-dir-${Math.random().toString(36).slice(2)}`;
    const dirGroup = document.createElement('div');
    dirGroup.className = 'fc-pastesp__group';
    const dirLegend = document.createElement('div');
    dirLegend.className = 'fc-pastesp__legend';
    dirLegend.textContent = t.seriesIn;
    const dirList = document.createElement('div');
    dirList.className = 'fc-pastesp__list';
    dirList.setAttribute('role', 'radiogroup');
    dirList.setAttribute('aria-label', dirLegend.textContent);
    const initialDirection = inferFillSeriesDirection(range);
    const directionOptions: Array<{ value: RibbonFillDirection; label: string }> = [
      { value: 'down', label: t.columns },
      { value: 'right', label: t.rows },
      { value: 'up', label: t.up },
      { value: 'left', label: t.left },
    ];
    for (const option of directionOptions) {
      dirList.appendChild(
        makeFillSeriesRadio(dirName, option.value, option.label, option.value === initialDirection),
      );
    }
    dirGroup.append(dirLegend, dirList);
    cols.appendChild(dirGroup);

    const modeName = `app-fill-series-mode-${Math.random().toString(36).slice(2)}`;
    const modeGroup = document.createElement('div');
    modeGroup.className = 'fc-pastesp__group';
    const modeLegend = document.createElement('div');
    modeLegend.className = 'fc-pastesp__legend';
    modeLegend.textContent = t.type;
    const modeList = document.createElement('div');
    modeList.className = 'fc-pastesp__list';
    modeList.setAttribute('role', 'radiogroup');
    modeList.setAttribute('aria-label', modeLegend.textContent);
    const modeOptions: Array<{ value: RibbonFillSeriesMode; label: string }> = [
      { value: 'auto', label: t.autoFill },
      { value: 'copy', label: t.copy },
      { value: 'days', label: t.day },
      { value: 'weekdays', label: t.weekday },
      { value: 'months', label: t.month },
      { value: 'years', label: t.year },
    ];
    for (const option of modeOptions) {
      modeList.appendChild(
        makeFillSeriesRadio(modeName, option.value, option.label, option.value === 'auto'),
      );
    }
    modeGroup.append(modeLegend, modeList);
    cols.appendChild(modeGroup);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: t.cancel,
      okLabel: t.ok,
    });
    let finish!: (
      value: { direction: RibbonFillDirection; mode: RibbonFillSeriesMode } | null,
    ) => void;
    const apply = (): void => {
      finish({
        direction: selectedFillSeriesRadio<RibbonFillDirection>(
          shell.overlay,
          dirName,
          initialDirection,
        ),
        mode: selectedFillSeriesRadio<RibbonFillSeriesMode>(shell.overlay, modeName, 'auto'),
      });
    };
    ({ finish } = installDialogLifecycle<{
      direction: RibbonFillDirection;
      mode: RibbonFillSeriesMode;
    } | null>({
      shell,
      resolve,
      onSubmit: apply,
      onCancel: () => null,
    }));
    cancelBtn.addEventListener('click', () => finish(null));
    okBtn.addEventListener('click', apply);
    mountDialog(shell, () => {
      dirList.querySelector<HTMLInputElement>('input[type="radio"]')?.focus();
    });
  });
};
