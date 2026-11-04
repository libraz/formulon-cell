import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachIterativeDialog } from '../../../src/interact/iterative-dialog.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

interface FakeWb {
  handle: WorkbookHandle;
  setIterative: ReturnType<typeof vi.fn>;
  setIterativeProgress: ReturnType<typeof vi.fn>;
  /** Toggle to false to simulate engines without iterative capability. */
  supports: { value: boolean };
}

interface IterativeSettings {
  enabled: boolean;
  maxIterations: number;
  maxChange: number;
}

const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-iterdlg');
const enableInput = (): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>('.fc-iterdlg [data-calc-option="iterative"]');
const numberInput = (): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>('.fc-iterdlg input[type="number"]');
const textInput = (): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>('.fc-iterdlg input[type="text"]');
const okBtn = (): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>('.fc-iterdlg .fc-iterdlg__btn--primary');
const cancelBtn = (): HTMLButtonElement | null =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-iterdlg .fc-iterdlg__btn')).find(
    (b) => !b.classList.contains('fc-iterdlg__btn--primary'),
  ) ?? null;
const status = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-iterdlg__status');

const makeFakeWb = (supports = true, stored: IterativeSettings | null = null): FakeWb => {
  const supportsRef = { value: supports };
  const setIterative = vi.fn(() => true);
  const setIterativeProgress = vi.fn(() => true);
  const handle = {
    get capabilities() {
      return { iterativeProgress: supportsRef.value } as never;
    },
    setIterative,
    setIterativeProgress,
    // Engines predating the readback have no such method at all, which is the
    // shape the other cases here exercise.
    ...(stored ? { getIterative: () => stored } : {}),
  } as unknown as WorkbookHandle;
  return { handle, setIterative, setIterativeProgress, supports: supportsRef };
};

describe('attachIterativeDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('open hydrates inputs from defaults (disabled, 100, 0.001) and disables numerics', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    expect(overlay()?.hidden).toBe(false);
    expect(enableInput()?.checked).toBe(false);
    expect(numberInput()?.value).toBe('100');
    expect(textInput()?.value).toBe('0.001');
    // When disabled, the numeric inputs are also disabled.
    expect(numberInput()?.disabled).toBe(true);
    expect(textInput()?.disabled).toBe(true);
    expect(numberInput()?.dataset.disabledReason).toBe(
      'この値を編集するには反復計算を有効にしてください。',
    );
    expect(textInput()?.getAttribute('aria-description')).toBe(
      'この値を編集するには反復計算を有効にしてください。',
    );
    handle.detach();
  });

  it('toggling the enable checkbox enables the numeric inputs', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const cb = enableInput();
    if (!cb) throw new Error('checkbox missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));
    expect(numberInput()?.disabled).toBe(false);
    expect(textInput()?.disabled).toBe(false);
    expect(numberInput()?.dataset.disabledReason).toBeUndefined();
    expect(textInput()?.getAttribute('aria-description')).toBeNull();
    handle.detach();
  });

  it('OK forwards the draft settings to wb.setIterative and wires a progress callback when enabled', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const cb = enableInput();
    if (!cb) throw new Error('checkbox missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));

    const num = numberInput();
    if (num) {
      num.value = '250';
      num.dispatchEvent(new Event('input', { bubbles: true }));
    }
    const tx = textInput();
    if (tx) {
      tx.value = '0.05';
      tx.dispatchEvent(new Event('input', { bubbles: true }));
    }
    okBtn()?.click();
    expect(fake.setIterative).toHaveBeenCalledTimes(1);
    expect(fake.setIterative).toHaveBeenCalledWith(true, 250, 0.05);
    // Enabled case wires a progress callback.
    expect(fake.setIterativeProgress).toHaveBeenCalledTimes(1);
    const arg = fake.setIterativeProgress.mock.calls[0]?.[0];
    expect(typeof arg).toBe('function');
    expect(overlay()?.hidden).toBe(true);
    handle.detach();
  });

  it('Enter acts as the default OK command', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const cb = enableInput();
    if (!cb) throw new Error('checkbox missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));

    overlay()?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(fake.setIterative).toHaveBeenCalledWith(true, 100, 0.001);
    expect(overlay()?.hidden).toBe(true);
    handle.detach();
  });

  it('numeric input clamps min iterations to 1', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const cb = enableInput();
    const num = numberInput();
    if (!cb || !num) throw new Error('controls missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));
    num.value = '0';
    num.dispatchEvent(new Event('input', { bubbles: true }));
    okBtn()?.click();
    // 0 → clamped up to 1.
    expect(fake.setIterative).toHaveBeenCalledWith(true, 1, 0.001);
    handle.detach();
  });

  it('Cancel closes without invoking setIterative', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const cb = enableInput();
    if (!cb) throw new Error('checkbox missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));
    cancelBtn()?.click();
    expect(fake.setIterative).not.toHaveBeenCalled();
    expect(overlay()?.hidden).toBe(true);
    handle.detach();
  });

  it('OK shows the unsupported message when wb.setIterative returns false', () => {
    const fake = makeFakeWb();
    fake.setIterative.mockReturnValueOnce(false);
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const cb = enableInput();
    if (!cb) throw new Error('checkbox missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));
    okBtn()?.click();
    expect(status()?.textContent).not.toBe('');
    // Dialog stays open on failure so the user can adjust.
    expect(overlay()?.hidden).toBe(false);
    handle.detach();
  });

  it('detach removes the overlay from the DOM', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.detach();
    expect(overlay()).toBeNull();
  });

  it('keeps Iterative Calculation on compact desktop dialog geometry', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/dialogs/iterative.css'), 'utf8');

    expect(css).toMatch(/\.fc-iterdlg__panel\s*\{[\s\S]*?width: 360px;/);
    expect(css).not.toMatch(/\.fc-iterdlg\s*\{/);
    expect(css).not.toMatch(/\.fc-iterdlg__(header|footer|btn)/);
    expect(css).not.toContain('box-shadow: 0 0 0 2px var(--fc-accent-soft');
  });

  it('wears the shared Format Cells dialog chrome', () => {
    const handle = attachIterativeDialog({ host, getWb: () => makeFakeWb().handle });
    const root = overlay();
    expect(root?.classList.contains('fc-fmtdlg')).toBe(true);
    expect(root?.querySelector('.fc-fmtdlg__panel')).not.toBeNull();
    expect(root?.querySelector('.fc-fmtdlg__header')).not.toBeNull();
    expect(root?.querySelector('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary')).not.toBeNull();
    handle.detach();
  });

  it('opens on the settings the workbook already carries', () => {
    const fake = makeFakeWb(true, { enabled: true, maxIterations: 250, maxChange: 0.05 });
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    expect(enableInput()?.checked).toBe(true);
    expect(numberInput()?.value).toBe('250');
    expect(textInput()?.value).toBe('0.05');
    expect(numberInput()?.disabled).toBe(false);
    handle.detach();
  });

  it('keeps its own defaults when the engine has no readback', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    expect(enableInput()?.checked).toBe(false);
    expect(numberInput()?.value).toBe('100');
    handle.detach();
  });

  it('parses maximum change with the live locale and keeps full input precision', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({
      host,
      getWb: () => fake.handle,
      getLocale: () => 'de-DE',
    });
    handle.open();
    const cb = enableInput();
    const input = textInput();
    if (!cb || !input) throw new Error('controls missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));
    input.value = '0,00000000000000012';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    okBtn()?.click();

    expect(fake.setIterative).toHaveBeenCalledWith(true, 100, 1.2e-16);
    handle.detach();
  });

  it('does not confirm while maximum change is invalid', () => {
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({
      host,
      getWb: () => fake.handle,
      getLocale: () => 'fr-FR',
    });
    handle.open();
    const cb = enableInput();
    const input = textInput();
    if (!cb || !input) throw new Error('controls missing');
    cb.checked = true;
    cb.dispatchEvent(new Event('change', { bubbles: true }));
    input.value = '0,1,2';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    okBtn()?.click();

    expect(fake.setIterative).not.toHaveBeenCalled();
    expect(input.validationMessage).not.toBe('');
    handle.detach();
  });

  it('uses the Mac Calculation pane wording on the Mac host', () => {
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    const fake = makeFakeWb();
    const handle = attachIterativeDialog({ host, getWb: () => fake.handle });
    handle.open();
    const visibleText = (selector: string): string[] =>
      [...document.querySelectorAll<HTMLElement>(selector)]
        .filter((el) => !el.hidden)
        .map((el) => el.textContent ?? '');
    expect(visibleText('.fc-iterdlg__row > span')).toEqual([
      '自動',
      'データ テーブル以外自動',
      '手動',
      'ブックの保存前に再計算を行う',
      '反復計算を使う',
      '最大反復回数:',
      '変化の最大値:',
      'R1C1 参照形式を使用する',
      '表示桁数で計算する',
      '1904 年から計算する',
    ]);
    expect(visibleText('.fc-iterdlg__section')).toEqual(['計算方法の設定', 'ブックの計算時']);
    expect(visibleText('.fc-iterdlg__label')).toEqual([]);
    expect(visibleText('.fc-iterdlg__note')).toEqual([]);
    handle.detach();
  });

  describe('on the Mac host', () => {
    // Excel for Mac's Calculation pane applies each change at once and has no OK / Cancel.
    const openMac = () => {
      host.classList.add('fc-host');
      host.dataset.fcPlatform = 'mac';
      const calls: string[] = [];
      const handle = {
        capabilities: { iterativeProgress: true, calcMode: true, calcSettings: true } as never,
        setIterative: vi.fn((...args: unknown[]) => {
          calls.push(`iter:${args.join(',')}`);
          return true;
        }),
        setIterativeProgress: vi.fn(() => true),
        setCalcMode: vi.fn((mode: number) => {
          calls.push(`mode:${mode}`);
          return true;
        }),
        setCalcSetting: vi.fn((key: string, value: unknown) => {
          calls.push(`${key}:${value}`);
          return true;
        }),
        recalcAuto: vi.fn(),
        cells: () => [],
      } as unknown as WorkbookHandle;
      const dialog = attachIterativeDialog({ host, getWb: () => handle });
      dialog.open();
      return { dialog, calls };
    };
    const change = (el: HTMLInputElement, checked?: boolean): void => {
      if (checked !== undefined) el.checked = checked;
      el.dispatchEvent(new Event('change', { bubbles: true }));
    };

    it('hides the OK and Cancel footer and shows a close control', () => {
      const { dialog } = openMac();
      expect(document.querySelector<HTMLElement>('.fc-iterdlg__footer')?.hidden).toBe(true);
      expect(document.querySelector<HTMLElement>('.fc-fmtdlg__close')?.hidden).toBe(false);
      dialog.detach();
    });

    it('applies the calculation mode as soon as it is picked', () => {
      const { dialog, calls } = openMac();
      const manual = document.querySelector<HTMLInputElement>(
        '.fc-iterdlg [data-calc-option="manual"]',
      );
      if (!manual) throw new Error('missing manual radio');
      change(manual, true);
      expect(calls).toContain('mode:1');
      dialog.detach();
    });

    it('applies iterative calculation and its limits without OK', () => {
      const { dialog, calls } = openMac();
      const enable = enableInput();
      if (!enable) throw new Error('missing checkbox');
      change(enable, true);
      expect(calls).toContain('iter:true,100,0.001');
      const max = numberInput();
      if (!max) throw new Error('missing input');
      max.value = '50';
      max.dispatchEvent(new Event('input', { bubbles: true }));
      change(max);
      expect(calls).toContain('iter:true,50,0.001');
      dialog.detach();
    });

    it('keeps OK semantics on the other profiles', () => {
      const fake = makeFakeWb();
      const dialog = attachIterativeDialog({ host, getWb: () => fake.handle });
      dialog.open();
      const cb = enableInput();
      if (!cb) throw new Error('missing checkbox');
      change(cb, true);
      expect(fake.setIterative).not.toHaveBeenCalled();
      dialog.detach();
    });
  });
});
