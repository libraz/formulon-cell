import { afterEach, describe, expect, it } from 'vitest';
import { attachRangePickerButton } from '../../../src/interact/range-picker-control.js';

const mount = () => {
  const dlg = document.createElement('div');
  dlg.className = 'fc-fmtdlg';
  const input = document.createElement('input');
  const other = document.createElement('input');
  dlg.append(input, other);
  document.body.append(dlg);
  const listeners = new Set<() => void>();
  let value = 'A1';
  const button = attachRangePickerButton(input, {
    label: 'pick',
    getValue: () => value,
    subscribeToRangeChanges: (l) => {
      listeners.add(l);
      return () => listeners.delete(l);
    },
  });
  return {
    dlg,
    input,
    other,
    button,
    listeners,
    pick: (v: string) => {
      value = v;
      for (const l of [...listeners]) l();
    },
  };
};

describe('range picker control', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('focusing the field lets grid picks reach it without collapsing the dialog', () => {
    const h = mount();
    h.input.focus();
    expect(h.dlg.classList.contains('fc-fmtdlg--range-armed')).toBe(true);
    expect(h.dlg.classList.contains('fc-fmtdlg--range-picking')).toBe(false);
    h.pick('B2');
    expect(h.input.value).toBe('B2');
  });

  it('returns focus to the field after a grid pick ends', async () => {
    const h = mount();
    h.input.focus();
    h.input.blur();
    h.pick('C3');
    document.body.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await Promise.resolve();
    expect(document.activeElement).toBe(h.input);
    expect(h.input.value).toBe('C3');
  });

  it('clicking the already-selected cell still writes the reference', async () => {
    const h = mount();
    h.input.focus();
    h.input.value = 'typed';
    h.input.blur();
    document.body.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await Promise.resolve();
    expect(h.input.value).toBe('A1');
    expect(document.activeElement).toBe(h.input);
  });

  it('a pointerup on a real control does not steal focus', async () => {
    const h = mount();
    h.input.focus();
    h.input.blur();
    h.other.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    const btn = document.createElement('button');
    document.body.append(btn);
    btn.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    await Promise.resolve();
    expect(document.activeElement).not.toBe(h.input);
  });

  it('keeps picking when focus leaves for the grid, stops for another dialog control', () => {
    const h = mount();
    h.input.focus();
    h.input.dispatchEvent(new FocusEvent('blur', { relatedTarget: document.body }));
    expect(h.dlg.classList.contains('fc-fmtdlg--range-armed')).toBe(true);
    h.input.dispatchEvent(new FocusEvent('blur', { relatedTarget: h.other }));
    expect(h.dlg.classList.contains('fc-fmtdlg--range-armed')).toBe(false);
    expect(h.listeners.size).toBe(0);
  });

  it('the picker button escalates an armed field to collapsed picking', () => {
    const h = mount();
    h.input.focus();
    h.button.click();
    expect(h.dlg.classList.contains('fc-fmtdlg--range-picking')).toBe(true);
    expect(h.dlg.classList.contains('fc-fmtdlg--range-armed')).toBe(false);
    h.button.click();
    expect(h.dlg.classList.contains('fc-fmtdlg--range-picking')).toBe(false);
  });
});
