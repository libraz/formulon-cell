import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachIterativeDialog } from '../../../src/interact/iterative-dialog.js';

interface FakeWb {
  handle: WorkbookHandle;
  setIterative: ReturnType<typeof vi.fn>;
  setIterativeProgress: ReturnType<typeof vi.fn>;
  /** Toggle to false to simulate engines without iterative capability. */
  supports: { value: boolean };
}

const overlay = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-iterdlg');
const enableInput = (): HTMLInputElement | null =>
  document.querySelector<HTMLInputElement>('.fc-iterdlg input[type="checkbox"]');
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

const makeFakeWb = (supports = true): FakeWb => {
  const supportsRef = { value: supports };
  const setIterative = vi.fn(() => true);
  const setIterativeProgress = vi.fn(() => true);
  const handle = {
    get capabilities() {
      return { iterativeProgress: supportsRef.value } as never;
    },
    setIterative,
    setIterativeProgress,
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
});
