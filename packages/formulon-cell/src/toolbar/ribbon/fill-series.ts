import type { Locale } from '../../i18n/strings.js';

// Fill Series subsystem: direction/mode types, the source-range collapse
// helper, the auto-direction heuristic, and the modal dialog that lets the
// user pick direction + series type. Kept self-contained so callers only own
// the imperative `runFillSeries` / `applyFillSeries` glue.

import type { Range } from '../../engine/types.js';
import { dictionaries, type Strings } from '../../i18n/strings.js';
import { createRadioOption } from '../dialogs/form-controls.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from '../dialogs/shell.js';

export type RibbonFillDirection = 'down' | 'right' | 'up' | 'left';
export type RibbonFillSeriesMode =
  | 'auto'
  | 'copy'
  | 'linear'
  | 'growth'
  | 'days'
  | 'weekdays'
  | 'months'
  | 'years';
export type FillSeriesDialogText = Strings['fillSeriesDialog'];

export const fillSeriesSourceRange = (range: Range, direction: RibbonFillDirection): Range => {
  if (direction === 'down') return { ...range, r1: range.r0 };
  if (direction === 'up') return { ...range, r0: range.r1 };
  if (direction === 'right') return { ...range, c1: range.c0 };
  return { ...range, c0: range.c1 };
};

export const inferFillSeriesDirection = (range: Range): RibbonFillDirection =>
  range.r1 > range.r0 ? 'down' : 'right';

const makeFillSeriesRadio = <T extends string>(
  name: string,
  value: T,
  label: string,
  checked: boolean,
): HTMLLabelElement => {
  return createRadioOption(document, { name, value, label, checked }).wrap;
};

const selectedFillSeriesRadio = <T extends string>(
  root: HTMLElement,
  name: string,
  fallback: T,
): T =>
  (root.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value as T | undefined) ??
  fallback;

export const showFillSeriesDialog = (
  range: Range,
  text: FillSeriesDialogText | Locale,
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

    const appendRadioColumn = <V extends string>(
      namePrefix: string,
      legendText: string,
      options: ReadonlyArray<{ value: V; label: string }>,
      initial: V,
    ): { name: string; list: HTMLDivElement } => {
      const name = `${namePrefix}-${Math.random().toString(36).slice(2)}`;
      const group = document.createElement('div');
      group.className = 'fc-pastesp__group';
      const legend = document.createElement('div');
      legend.className = 'fc-pastesp__legend';
      legend.textContent = legendText;
      const list = document.createElement('div');
      list.className = 'fc-pastesp__list';
      list.setAttribute('role', 'radiogroup');
      list.setAttribute('aria-label', legendText);
      for (const option of options) {
        list.appendChild(
          makeFillSeriesRadio(name, option.value, option.label, option.value === initial),
        );
      }
      group.append(legend, list);
      cols.appendChild(group);
      return { name, list };
    };

    const initialDirection = inferFillSeriesDirection(range);
    const { name: dirName, list: dirList } = appendRadioColumn<RibbonFillDirection>(
      'app-fill-series-dir',
      t.seriesIn,
      [
        { value: 'down', label: t.columns },
        { value: 'right', label: t.rows },
        { value: 'up', label: t.up },
        { value: 'left', label: t.left },
      ],
      initialDirection,
    );
    const { name: modeName } = appendRadioColumn<RibbonFillSeriesMode>(
      'app-fill-series-mode',
      t.type,
      [
        { value: 'auto', label: t.autoFill },
        { value: 'copy', label: t.copy },
        { value: 'linear', label: t.linear },
        { value: 'growth', label: t.growth },
        { value: 'days', label: t.day },
        { value: 'weekdays', label: t.weekday },
        { value: 'months', label: t.month },
        { value: 'years', label: t.year },
      ],
      'auto',
    );

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
