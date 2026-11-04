import { afterEach, describe, expect, it, vi } from 'vitest';
import { enhanceCustomSelect } from '../../../src/interact/custom-select.js';

afterEach(() => {
  vi.restoreAllMocks();
  document.body.replaceChildren();
});

const make = (): HTMLSelectElement => {
  const select = document.createElement('select');
  select.innerHTML = '<option value="a">A</option><option value="b">B</option>';
  document.body.append(select);
  return select;
};

const net = (add: ReturnType<typeof vi.spyOn>, remove: ReturnType<typeof vi.spyOn>): number => {
  const count = (spy: ReturnType<typeof vi.spyOn>): number =>
    spy.mock.calls.filter((call: unknown[]) => String(call[0]).startsWith('pointer')).length;
  return count(add) - count(remove);
};

describe('enhanceCustomSelect document listeners', () => {
  it('holds document pointer listeners only while the list is open', () => {
    const add = vi.spyOn(document, 'addEventListener');
    const remove = vi.spyOn(document, 'removeEventListener');
    const select = make();
    const handle = enhanceCustomSelect(select);
    expect(net(add, remove)).toBe(0);

    const button = select.parentElement?.querySelector('button') as HTMLButtonElement;
    button.click();
    expect(net(add, remove)).toBe(3);
    button.click();
    expect(net(add, remove)).toBe(0);

    button.click();
    expect(net(add, remove)).toBe(3);
    handle?.dispose();
    expect(net(add, remove)).toBe(0);
  });

  it('still closes on a pointerdown outside', () => {
    const select = make();
    enhanceCustomSelect(select);
    const button = select.parentElement?.querySelector('button') as HTMLButtonElement;
    button.click();
    expect(button.getAttribute('aria-expanded')).toBe('true');
    document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    expect(button.getAttribute('aria-expanded')).toBe('false');
  });
});
