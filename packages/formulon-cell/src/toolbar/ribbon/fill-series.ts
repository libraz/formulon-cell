// Fill Series subsystem: direction/mode types, the source-range collapse
// helper, the auto-direction heuristic, and the modal dialog that lets the
// user pick direction + series type. Kept self-contained so main.ts only
// owns the imperative `runFillSeries` / `applyFillSeries` glue.

import type { Range } from '@libraz/formulon-cell';

export type RibbonFillDirection = 'down' | 'right' | 'up' | 'left';
export type RibbonFillSeriesMode = 'auto' | 'copy' | 'days' | 'weekdays' | 'months' | 'years';

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
  ribbonLang: 'ja' | 'en',
): Promise<{ direction: RibbonFillDirection; mode: RibbonFillSeriesMode } | null> => {
  return new Promise((resolve) => {
    const ja = ribbonLang === 'ja';
    const title = ja ? '連続データ' : 'Series';
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel fc-pastesp__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body fc-pastesp__body';
    panel.appendChild(body);

    const cols = document.createElement('div');
    cols.className = 'fc-pastesp__cols';
    body.appendChild(cols);

    const dirName = `app-fill-series-dir-${Math.random().toString(36).slice(2)}`;
    const dirGroup = document.createElement('div');
    dirGroup.className = 'fc-pastesp__group';
    const dirLegend = document.createElement('div');
    dirLegend.className = 'fc-pastesp__legend';
    dirLegend.textContent = ja ? '範囲' : 'Series in';
    const dirList = document.createElement('div');
    dirList.className = 'fc-pastesp__list';
    dirList.setAttribute('role', 'radiogroup');
    dirList.setAttribute('aria-label', dirLegend.textContent);
    const initialDirection = inferFillSeriesDirection(range);
    const directionOptions: Array<{ value: RibbonFillDirection; label: string }> = [
      { value: 'down', label: ja ? '列' : 'Columns' },
      { value: 'right', label: ja ? '行' : 'Rows' },
      { value: 'up', label: ja ? '上方向' : 'Up' },
      { value: 'left', label: ja ? '左方向' : 'Left' },
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
    modeLegend.textContent = ja ? '種類' : 'Type';
    const modeList = document.createElement('div');
    modeList.className = 'fc-pastesp__list';
    modeList.setAttribute('role', 'radiogroup');
    modeList.setAttribute('aria-label', modeLegend.textContent);
    const modeOptions: Array<{ value: RibbonFillSeriesMode; label: string }> = [
      { value: 'auto', label: ja ? 'オートフィル' : 'AutoFill' },
      { value: 'copy', label: ja ? 'コピー' : 'Copy' },
      { value: 'days', label: ja ? '日' : 'Day' },
      { value: 'weekdays', label: ja ? '週日' : 'Weekday' },
      { value: 'months', label: ja ? '月' : 'Month' },
      { value: 'years', label: ja ? '年' : 'Year' },
    ];
    for (const option of modeOptions) {
      modeList.appendChild(
        makeFillSeriesRadio(modeName, option.value, option.label, option.value === 'auto'),
      );
    }
    modeGroup.append(modeLegend, modeList);
    cols.appendChild(modeGroup);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'fc-fmtdlg__btn';
    cancelBtn.textContent = ja ? 'キャンセル' : 'Cancel';
    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    okBtn.textContent = 'OK';
    footer.append(cancelBtn, okBtn);

    let done = false;
    const finish = (
      value: { direction: RibbonFillDirection; mode: RibbonFillSeriesMode } | null,
    ): void => {
      if (done) return;
      done = true;
      overlay.removeEventListener('keydown', onKey);
      overlay.remove();
      opener?.focus({ preventScroll: true });
      resolve(value);
    };
    const apply = (): void => {
      finish({
        direction: selectedFillSeriesRadio<RibbonFillDirection>(overlay, dirName, initialDirection),
        mode: selectedFillSeriesRadio<RibbonFillSeriesMode>(overlay, modeName, 'auto'),
      });
    };
    const onKey = (event: KeyboardEvent): void => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        event.preventDefault();
        finish(null);
      } else if (event.key === 'Enter') {
        event.preventDefault();
        apply();
      }
    };
    cancelBtn.addEventListener('click', () => finish(null));
    okBtn.addEventListener('click', apply);
    overlay.addEventListener('keydown', onKey);
    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) finish(null);
    });
    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      dirList.querySelector<HTMLInputElement>('input[type="radio"]')?.focus();
    });
  });
};
