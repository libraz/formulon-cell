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
