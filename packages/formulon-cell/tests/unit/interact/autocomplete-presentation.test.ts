import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { attachAutocomplete } from '../../../src/interact/autocomplete.js';

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('autocomplete presentation and lifecycle', () => {
  const attached: { input: HTMLTextAreaElement; detach: () => void }[] = [];

  afterEach(() => {
    for (const { input, detach } of attached.splice(0)) {
      detach();
      input.remove();
    }
    vi.restoreAllMocks();
  });

  const setup = (value: string, names: readonly string[] = ['SUM', 'SUMIF']) => {
    const input = document.createElement('textarea');
    input.value = value;
    document.body.append(input);
    input.setSelectionRange(value.length, value.length);
    const handle = attachAutocomplete({ input, getFunctionNames: () => names });
    attached.push({ input, detach: () => handle.detach() });
    handle.refresh();
    return { input, handle };
  };

  it('keeps the listbox limited to options and exposes active details in a sibling ScreenTip', () => {
    const { input } = setup('=SU');
    const panel = document.querySelector<HTMLElement>('.fc-autocomplete');
    const listbox = panel?.querySelector<HTMLElement>('[role="listbox"]');
    const options = [...(listbox?.children ?? [])];
    const tip = panel?.querySelector<HTMLElement>('[role="tooltip"]');

    expect(panel?.getAttribute('role')).toBeNull();
    expect(listbox?.id).toBe(input.getAttribute('aria-controls'));
    expect(options.length).toBeGreaterThan(0);
    expect(options.every((option) => option.getAttribute('role') === 'option')).toBe(true);
    expect(tip?.textContent).toContain('SUM(');
    expect(listbox?.contains(tip ?? null)).toBe(false);
    expect(options[0]?.getAttribute('aria-describedby')).toBe(tip?.id);
    expect(input.getAttribute('aria-activedescendant')).toBe(options[0]?.id);
  });

  it('keeps option nodes stable while highlighting and updates only the active identity', () => {
    const { input, handle } = setup('=S');
    const first = document.querySelector('.fc-autocomplete__item');
    const second = document.querySelectorAll('.fc-autocomplete__item')[1];

    handle.move(1);

    expect(document.querySelectorAll('.fc-autocomplete__item')[0]).toBe(first);
    expect(document.querySelectorAll('.fc-autocomplete__item')[1]).toBe(second);
    expect(first?.getAttribute('aria-selected')).toBe('false');
    expect(second?.getAttribute('aria-selected')).toBe('true');

    const activeId = input.getAttribute('aria-activedescendant');
    input.value = '=SU';
    input.setSelectionRange(input.value.length, input.value.length);
    handle.refresh();
    expect(input.getAttribute('aria-activedescendant')).toBe(activeId);
    expect(document.querySelectorAll('.fc-autocomplete__item')[1]).toBe(second);
  });

  it('keeps unique option IDs as candidates are added and removed across refreshes', () => {
    const input = document.createElement('textarea');
    input.value = '=TEST';
    document.body.append(input);
    input.setSelectionRange(input.value.length, input.value.length);
    let names = ['TESTA', 'TESTB', 'TESTC'];
    const handle = attachAutocomplete({ input, getFunctionNames: () => names });
    attached.push({ input, detach: () => handle.detach() });
    handle.refresh();
    const optionByName = (): Map<string, HTMLButtonElement> =>
      new Map(
        [...document.querySelectorAll<HTMLButtonElement>('.fc-autocomplete__item')].map(
          (option) => [option.querySelector('.fc-autocomplete__name')?.textContent ?? '', option],
        ),
      );
    const original = optionByName();
    const testB = original.get('TESTB');
    const testC = original.get('TESTC');

    names = ['TESTB', 'TESTC', 'TESTD'];
    handle.refresh();
    expect(optionByName().get('TESTB')).toBe(testB);
    expect(optionByName().get('TESTC')).toBe(testC);

    names = ['TESTB', 'TESTC', 'TESTD', 'TESTE'];
    handle.refresh();
    const options = [...document.querySelectorAll<HTMLButtonElement>('.fc-autocomplete__item')];
    const ids = options.map((option) => option.id);
    expect(options).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
    expect(optionByName().get('TESTB')).toBe(testB);
    expect(optionByName().get('TESTC')).toBe(testC);
    expect(input.getAttribute('aria-activedescendant')).toBeTruthy();
    expect(
      options.some((option) => option.id === input.getAttribute('aria-activedescendant')),
    ).toBe(true);
  });

  it('aligns the compact ScreenTip with the active row inside the scrolled list', () => {
    const { handle } = setup('=S');
    const list = document.querySelector<HTMLElement>('.fc-autocomplete__list');
    const second = document.querySelector<HTMLElement>('.fc-autocomplete__item:nth-child(2)');
    const help = document.querySelector<HTMLElement>('.fc-autocomplete__help');
    expect(list).not.toBeNull();
    expect(second).not.toBeNull();
    expect(help).not.toBeNull();
    Object.defineProperty(list, 'clientHeight', { configurable: true, value: 200 });
    Object.defineProperty(second, 'offsetTop', { configurable: true, value: 72 });
    Object.defineProperty(help, 'offsetHeight', { configurable: true, value: 56 });
    if (list) list.scrollTop = 24;

    handle.move(1);

    expect(help?.style.marginTop).toBe('48px');
  });

  it('keeps the candidate-list origin fixed as active metadata appears and disappears', () => {
    const input = document.createElement('textarea');
    input.value = '=TEST';
    document.body.append(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attachAutocomplete({
      input,
      getFunctionNames: () => ['TESTA', 'TESTB'],
      getFunctionDetails: (name) =>
        name === 'TESTA' ? { signature: 'TESTA(value)', description: 'Test details.' } : undefined,
    });
    attached.push({ input, detach: () => handle.detach() });
    handle.refresh();

    const panel = document.querySelector<HTMLElement>('.fc-autocomplete');
    const list = panel?.querySelector<HTMLElement>('.fc-autocomplete__list');
    const help = panel?.querySelector<HTMLElement>('.fc-autocomplete__help');
    expect(panel).not.toBeNull();
    expect(list).not.toBeNull();
    expect(help?.textContent).toContain('TESTA(value)');
    expect(panel?.classList.contains('fc-autocomplete--has-help')).toBe(true);

    handle.move(1);
    expect(help?.textContent).toBe('');
    expect(panel?.classList.contains('fc-autocomplete--has-help')).toBe(false);
    expect(panel?.querySelector('.fc-autocomplete__list')).toBe(list);

    handle.move(-1);
    expect(help?.textContent).toContain('TESTA(value)');
    expect(panel?.classList.contains('fc-autocomplete--has-help')).toBe(true);
    expect(panel?.querySelector('.fc-autocomplete__list')).toBe(list);

    const css = readFileSync(
      resolve(packageRoot, 'src/styles/core/app/popups/formula-helpers.css'),
      'utf8',
    );
    expect(css).toMatch(/\.fc-autocomplete\s*\{[^}]*grid-template-columns:\s*220px 224px;/);
    expect(css).not.toMatch(/\.fc-autocomplete--has-help\s*\{[^}]*grid-template-columns:/);
    expect(css).toMatch(/\.fc-autocomplete__help:empty\s*\{[^}]*visibility:\s*hidden;/);
    expect(css).not.toMatch(/\.fc-autocomplete__help:empty\s*\{[^}]*display:\s*none;/);
    expect(css).toMatch(
      /\.fc-autocomplete\.fc-autocomplete--narrow \.fc-autocomplete__help\s*\{[^}]*height:\s*72px;[^}]*max-height:\s*72px;/,
    );
  });

  it('stacks help when a scaled portal halves the available CSS width', () => {
    const input = document.createElement('textarea');
    input.value = '=S';
    document.body.append(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attachAutocomplete({ input, getFunctionNames: () => ['SUM'] });
    attached.push({ input, detach: () => handle.detach() });
    handle.refresh();

    const panel = document.querySelector<HTMLElement>('.fc-autocomplete');
    expect(panel).not.toBeNull();
    if (!panel) return;
    Object.defineProperty(panel, 'offsetWidth', { configurable: true, value: 452 });
    Object.defineProperty(panel, 'offsetHeight', { configurable: true, value: 200 });
    vi.spyOn(panel, 'getBoundingClientRect').mockImplementation(
      () =>
        ({
          x: 20,
          y: 20,
          left: 20,
          top: 20,
          right: 924,
          bottom: 420,
          width: 904,
          height: 400,
          toJSON: () => ({}),
        }) as DOMRect,
    );
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(840);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(620);

    handle.refresh();

    expect(panel.classList.contains('fc-autocomplete--narrow')).toBe(true);
    expect(panel.style.maxWidth).toBe('414px');
  });

  it('uses single mousedown to highlight and double-click to accept', () => {
    const { input } = setup('=SU');
    const options = [...document.querySelectorAll<HTMLButtonElement>('.fc-autocomplete__item')];
    const event = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    options[1]?.dispatchEvent(event);

    expect(event.defaultPrevented).toBe(true);
    expect(input.value).toBe('=SU');
    expect(options[1]?.getAttribute('aria-selected')).toBe('true');

    options[1]?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(input.value).toBe('=SUMIF(');
  });

  it('sorts and deduplicates every prefix match while paging past the visible rows', () => {
    const names = [
      'SUM09',
      'SUM02',
      'SUM01',
      'SUM03',
      'SUM04',
      'SUM05',
      'SUM06',
      'SUM07',
      'SUM08',
      'sum02',
    ];
    const { input, handle } = setup('=SU', names);
    const options = [...document.querySelectorAll<HTMLElement>('.fc-autocomplete__item')];

    expect(
      options.map((option) => option.querySelector('.fc-autocomplete__name')?.textContent),
    ).toEqual(['SUM01', 'SUM02', 'SUM03', 'SUM04', 'SUM05', 'SUM06', 'SUM07', 'SUM08', 'SUM09']);
    expect(handle.navigate).toBeTypeOf('function');
    handle.navigate?.('last');
    expect(input.getAttribute('aria-activedescendant')).toBe(options[8]?.id);
    handle.navigate?.('first');
    handle.navigate?.('pageDown');
    expect(input.getAttribute('aria-activedescendant')).toBe(options[8]?.id);
    handle.navigate?.('pageUp');
    expect(input.getAttribute('aria-activedescendant')).toBe(options[0]?.id);
    handle.navigate?.('last');
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=SUM09(');
  });

  it('dismisses the same text and caret until editing or explicit toggle', () => {
    const { input, handle } = setup('=SU');

    expect(handle.dismiss).toBeTypeOf('function');
    handle.dismiss?.();
    expect(handle.isOpen()).toBe(false);
    handle.refresh();
    expect(handle.isOpen()).toBe(false);

    handle.toggle?.();
    expect(handle.isOpen()).toBe(true);
    handle.dismiss?.();
    input.value = '=SUM';
    input.setSelectionRange(4, 4);
    handle.refresh();
    expect(handle.isOpen()).toBe(true);
  });

  it('suppresses immediate reopen after accepting a candidate', () => {
    const { input, handle } = setup('=SU');
    expect(handle.acceptHighlighted()).toBe(true);

    handle.refresh();

    expect(input.value).toBe('=SUM(');
    expect(handle.isOpen()).toBe(false);
  });

  it('hides during composition and refreshes once composition ends', () => {
    const { input, handle } = setup('=SU');

    expect(handle.setComposing).toBeTypeOf('function');
    handle.setComposing?.(true);
    expect(handle.isOpen()).toBe(false);
    handle.refresh();
    expect(handle.isOpen()).toBe(false);

    handle.setComposing?.(false);
    expect(handle.isOpen()).toBe(true);
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    expect(handle.isOpen()).toBe(false);
    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
    expect(handle.isOpen()).toBe(true);
  });

  it('uses supplied function ScreenTip details and leaves unknown labels empty', () => {
    const input = document.createElement('textarea');
    input.value = '=SUB';
    document.body.append(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attachAutocomplete({
      input,
      getFunctionNames: () => ['SUBTOTAL', 'SUBSTITUTE'],
      getFunctionDetails: (name) =>
        name === 'SUBTOTAL'
          ? {
              signature: 'SUBTOTAL(function_num, ref1, [ref2], ...)',
              description: 'Returns a subtotal.',
            }
          : undefined,
    });
    attached.push({ input, detach: () => handle.detach() });

    handle.refresh();

    const tip = document.querySelector<HTMLElement>('.fc-autocomplete__help');
    handle.navigate?.('last');
    expect(tip?.textContent).toContain('SUBTOTAL(function_num, ref1, [ref2], ...)');
    expect(tip?.textContent).toContain('Returns a subtotal.');
    handle.navigate?.('first');
    expect(tip?.textContent).not.toContain('Custom function');
  });
});
