import { afterEach, describe, expect, it, vi } from 'vitest';
import { attachAutocomplete } from '../../../src/interact/autocomplete.js';

afterEach(() => {
  document.body.replaceChildren();
});

describe('attachAutocomplete highlight scrolling', () => {
  it('scrolls for keyboard moves but not for pointer hover', () => {
    const input = document.createElement('textarea');
    input.value = '=S';
    document.body.append(input);
    input.setSelectionRange(2, 2);
    const handle = attachAutocomplete({ input, getFunctionNames: () => ['SUM', 'SUMIF', 'SIN'] });
    handle.refresh();

    const options = [...document.querySelectorAll<HTMLElement>('[role="option"]')];
    expect(options.length).toBeGreaterThan(1);
    const spies = options.map((o) => {
      const fn = vi.fn();
      o.scrollIntoView = fn;
      return fn;
    });

    options[1]?.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }));
    expect(options[1]?.classList.contains('fc-autocomplete__item--active')).toBe(true);
    for (const s of spies) expect(s).not.toHaveBeenCalled();

    handle.move(1);
    const called = spies.reduce((n, s) => n + s.mock.calls.length, 0);
    expect(called).toBe(1);
    handle.detach();
  });
});
