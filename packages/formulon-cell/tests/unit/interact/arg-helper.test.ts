import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachArgHelper } from '../../../src/interact/arg-helper.js';

const tooltip = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-arghelper');
const args = (): HTMLElement[] =>
  Array.from(document.querySelectorAll<HTMLElement>('.fc-arghelper__arg'));
const activeArg = (): HTMLElement | undefined =>
  args().find((a) => a.classList.contains('fc-arghelper__arg--active'));

const setText = (input: HTMLInputElement, text: string, caret = text.length): void => {
  input.value = text;
  input.setSelectionRange(caret, caret);
};

describe('attachArgHelper', () => {
  let input: HTMLInputElement;

  beforeEach(() => {
    input = document.createElement('input');
    input.type = 'text';
    document.body.appendChild(input);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('shows the SUM signature with number1 active when caret sits just after the open paren', () => {
    setText(input, '=SUM(');
    const handle = attachArgHelper({ input });
    handle.refresh();
    expect(tooltip()).not.toBeNull();
    expect(tooltip()?.getAttribute('role')).toBe('tooltip');
    expect(tooltip()?.id).toMatch(/^fc-arghelper-/);
    expect(input.getAttribute('aria-describedby')).toBe(tooltip()?.id);
    const list = args().map((a) => a.textContent);
    expect(list).toEqual(['number1', '[number2]', '...']);
    expect(activeArg()?.textContent).toBe('number1');
    expect(activeArg()?.getAttribute('aria-current')).toBe('true');
    handle.detach();
  });

  it('advances the active argument across each top-level comma', () => {
    setText(input, '=IF(A1>0,');
    const handle = attachArgHelper({ input });
    handle.refresh();
    // After the first comma, value_if_true is the active arg.
    expect(activeArg()?.textContent).toBe('value_if_true');
    setText(input, '=IF(A1>0,1,');
    handle.refresh();
    expect(activeArg()?.textContent).toBe('[value_if_false]');
    handle.detach();
  });

  it('uses the innermost function when the caret is inside a nested call', () => {
    setText(input, '=IF(SUM(');
    const handle = attachArgHelper({ input });
    handle.refresh();
    // Innermost = SUM, so its argument list is rendered.
    const list = args().map((a) => a.textContent);
    expect(list).toEqual(['number1', '[number2]', '...']);
    expect(activeArg()?.textContent).toBe('number1');
    handle.detach();
  });

  it('hides the tooltip when the caret moves past the closing paren', () => {
    setText(input, '=SUM(');
    const handle = attachArgHelper({ input });
    handle.refresh();
    expect(tooltip()).not.toBeNull();
    setText(input, '=SUM(1)');
    handle.refresh();
    expect(tooltip()).toBeNull();
    expect(input.hasAttribute('aria-describedby')).toBe(false);
    handle.detach();
  });

  it('hides the tooltip for non-formula text', () => {
    setText(input, 'plain text not a formula');
    const handle = attachArgHelper({ input });
    handle.refresh();
    expect(tooltip()).toBeNull();
    handle.detach();
  });

  it('hides the tooltip for a function name without an open paren', () => {
    setText(input, '=SUM');
    const handle = attachArgHelper({ input });
    handle.refresh();
    expect(tooltip()).toBeNull();
    handle.detach();
  });

  it('detach removes the tooltip element from the DOM', () => {
    setText(input, '=SUM(');
    const handle = attachArgHelper({ input });
    handle.refresh();
    expect(tooltip()).not.toBeNull();
    handle.detach();
    expect(tooltip()).toBeNull();
    expect(input.hasAttribute('aria-describedby')).toBe(false);
  });
});
