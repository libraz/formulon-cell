import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { enhanceCustomSelect } from '../../../src/interact/custom-select.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

function requiredElement<T extends Element>(element: T | null): T {
  if (!element) throw new Error('Expected test element to exist');
  return element;
}

describe('enhanceCustomSelect', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('projects disabled option reasons through the shared a11y helper', () => {
    const select = document.createElement('select');
    select.setAttribute('aria-label', 'Calculation mode');
    select.innerHTML = `
      <option value="auto">Automatic</option>
      <option value="tables" disabled data-disabled-reason="Requires engine support">Data tables</option>
    `;
    document.body.appendChild(select);

    const handle = enhanceCustomSelect(select);
    document.querySelector<HTMLButtonElement>('.fc-select__button')?.click();

    const disabled = document.querySelector<HTMLButtonElement>(
      '.fc-select__option[data-value="tables"]',
    );
    const enabled = document.querySelector<HTMLButtonElement>(
      '.fc-select__option[data-value="auto"]',
    );

    expect(disabled?.disabled).toBe(true);
    expect(disabled?.dataset.disabledReason).toBe('Requires engine support');
    expect(disabled?.getAttribute('aria-description')).toBe('Requires engine support');
    expect(disabled?.title).toBe('Requires engine support');
    expect(enabled?.dataset.disabledReason).toBeUndefined();
    expect(enabled?.getAttribute('aria-description')).toBeNull();

    handle?.dispose();
  });

  it('moves the highlight to Home and End targets and commits the highlighted option', () => {
    const select = document.createElement('select');
    select.innerHTML = `
      <option value="first">First</option>
      <option value="middle">Middle</option>
      <option value="last">Last</option>
    `;
    select.selectedIndex = 1;
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = document.querySelector<HTMLButtonElement>('.fc-select__button');
    const activeValue = (): string | undefined =>
      document.querySelector<HTMLElement>('.fc-select__option--active')?.dataset.value;

    button?.click();
    expect(activeValue()).toBe('middle');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(activeValue()).toBe('first');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(select.value).toBe('first');

    button?.click();
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(activeValue()).toBe('last');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(select.value).toBe('last');
    handle?.dispose();
  });

  it('lets a closed custom select pass Escape to its parent', () => {
    const select = document.createElement('select');
    select.innerHTML = '<option value="one">One</option>';
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = document.querySelector<HTMLButtonElement>('.fc-select__button');
    let parentSawEscape = false;
    document.body.addEventListener(
      'keydown',
      (event) => {
        parentSawEscape = event.key === 'Escape';
      },
      { once: true },
    );

    const closedEscapeEvent = new KeyboardEvent('keydown', {
      key: 'Escape',
      bubbles: true,
      cancelable: true,
    });
    button?.dispatchEvent(closedEscapeEvent);

    expect(parentSawEscape).toBe(true);
    expect(closedEscapeEvent.defaultPrevented).toBe(false);
    handle?.dispose();
  });

  it('mirrors the native disabled state and disabled reason on its trigger', async () => {
    const select = document.createElement('select');
    select.disabled = true;
    select.dataset.disabledReason = 'Unavailable';
    select.innerHTML = '<option value="one">One</option><option value="two">Two</option>';
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = document.querySelector<HTMLButtonElement>('.fc-select__button');
    expect(button?.disabled).toBe(true);
    expect(button?.dataset.disabledReason).toBe('Unavailable');
    const disabledValue = select.value;
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(select.value).toBe(disabledValue);

    select.disabled = false;
    delete select.dataset.disabledReason;
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(button?.disabled).toBe(false);
    expect(button?.dataset.disabledReason).toBeUndefined();
    handle?.dispose();
  });

  it('skips disabled options and disabled optgroups while moving the highlight', () => {
    const select = document.createElement('select');
    select.innerHTML = `
      <option value="start">Start</option>
      <option value="disabled" disabled>Disabled</option>
      <optgroup label="Unavailable" disabled>
        <option value="group-one">Group one</option>
        <option value="group-two">Group two</option>
      </optgroup>
      <option value="end">End</option>
    `;
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = document.querySelector<HTMLButtonElement>('.fc-select__button');
    const row = (value: string): HTMLButtonElement | null =>
      document.querySelector<HTMLButtonElement>(`.fc-select__option[data-value="${value}"]`);
    const activeValue = (): string | undefined =>
      document.querySelector<HTMLElement>('.fc-select__option--active')?.dataset.value;

    expect(row('disabled')?.disabled).toBe(true);
    expect(row('group-one')?.disabled).toBe(true);
    button?.click();
    row('disabled')?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(select.value).toBe('start');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(activeValue()).toBe('end');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(activeValue()).toBe('start');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
    expect(activeValue()).toBe('end');
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(select.value).toBe('end');
    handle?.dispose();
  });

  it('keeps empty and all-disabled option lists bounded and inert', () => {
    const disabledSelect = document.createElement('select');
    disabledSelect.innerHTML = `
      <option value="one" disabled>One</option>
      <option value="two" disabled>Two</option>
    `;
    document.body.appendChild(disabledSelect);
    const disabledHandle = enhanceCustomSelect(disabledSelect);
    const disabledButton = requiredElement(disabledSelect.nextElementSibling) as HTMLButtonElement;
    const disabledList = requiredElement(
      requiredElement(disabledSelect.parentElement).querySelector<HTMLElement>('.fc-select__list'),
    );
    expect(disabledButton).toBeInstanceOf(HTMLButtonElement);
    const disabledValue = disabledSelect.value;
    disabledButton.click();
    expect(disabledList.hidden).toBe(false);
    for (const key of ['Home', 'End', 'ArrowDown', 'ArrowUp', 'Enter']) {
      disabledButton.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
    }
    expect(disabledSelect.value).toBe(disabledValue);
    expect(disabledList.querySelector('.fc-select__option--active')).toBeNull();

    const emptySelect = document.createElement('select');
    document.body.appendChild(emptySelect);
    const emptyHandle = enhanceCustomSelect(emptySelect);
    const emptyButton = requiredElement(emptySelect.nextElementSibling) as HTMLButtonElement;
    const emptyList = requiredElement(
      requiredElement(emptySelect.parentElement).querySelector<HTMLElement>('.fc-select__list'),
    );
    expect(emptyButton).toBeInstanceOf(HTMLButtonElement);
    emptyButton.click();
    expect(emptyList.hidden).toBe(false);
    for (const key of ['Home', 'End', 'ArrowDown', 'ArrowUp', 'Enter']) {
      emptyButton.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
    }
    expect(emptySelect.value).toBe('');
    expect(emptyList.querySelector('[role="option"]')).toBeNull();
    disabledHandle?.dispose();
    emptyHandle?.dispose();
  });

  it('closes on Tab and focusout without committing the highlighted option', () => {
    const select = document.createElement('select');
    select.innerHTML = '<option value="one">One</option><option value="two">Two</option>';
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = document.querySelector<HTMLButtonElement>('.fc-select__button');
    const list = document.querySelector<HTMLElement>('.fc-select__list');
    let changes = 0;
    select.addEventListener('change', () => {
      changes += 1;
    });
    const outside = document.createElement('button');
    document.body.appendChild(outside);

    button?.click();
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
    button?.dispatchEvent(tab);
    expect(list?.hidden).toBe(true);
    expect(tab.defaultPrevented).toBe(false);
    expect(select.value).toBe('one');
    expect(changes).toBe(0);

    button?.click();
    button?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    button?.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: outside }));
    expect(list?.hidden).toBe(true);
    expect(select.value).toBe('one');
    expect(changes).toBe(0);
    outside.remove();
    handle?.dispose();
  });

  it('keeps an option click alive when WebKit reports null relatedTarget on focusout', () => {
    const select = document.createElement('select');
    select.innerHTML = '<option value="a">A</option><option value="d">D</option>';
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = select.nextElementSibling as HTMLButtonElement;
    const wrapper = requiredElement(select.parentElement);
    const list = requiredElement(wrapper.querySelector<HTMLElement>('.fc-select__list'));
    let changes = 0;
    select.addEventListener('change', () => {
      changes += 1;
    });
    let wrapperSawPointerDown = false;
    wrapper.addEventListener('pointerdown', () => {
      wrapperSawPointerDown = true;
    });

    button.focus();
    button.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(list.hidden).toBe(false);
    const option = requiredElement(
      list.querySelector<HTMLButtonElement>('.fc-select__option[data-value="d"]'),
    );
    expect(option.isConnected).toBe(true);
    option.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    expect(wrapperSawPointerDown).toBe(true);
    button.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
    expect(list.hidden).toBe(false);
    option.dispatchEvent(new Event('pointerup', { bubbles: true }));
    option.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(select.value).toBe('d');
    expect(changes).toBe(1);
    handle?.dispose();
  });

  it('clears pointer interaction state after cancel and window blur', () => {
    const select = document.createElement('select');
    select.innerHTML = '<option value="a">A</option><option value="d">D</option>';
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = requiredElement(select.nextElementSibling) as HTMLButtonElement;
    const list = requiredElement(
      requiredElement(select.parentElement).querySelector<HTMLElement>('.fc-select__list'),
    );

    button.click();
    button.focus();
    button.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    button.dispatchEvent(new Event('pointercancel', { bubbles: true }));
    button.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
    expect(list.hidden).toBe(true);

    button.click();
    button.focus();
    button.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    window.dispatchEvent(new Event('blur'));
    button.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
    expect(list.hidden).toBe(true);
    handle?.dispose();
  });

  it('closes if a pointer interaction starts inside and ends outside the custom select', () => {
    const select = document.createElement('select');
    select.innerHTML = '<option value="a">A</option><option value="d">D</option>';
    document.body.appendChild(select);
    const handle = enhanceCustomSelect(select);
    const button = requiredElement(select.nextElementSibling) as HTMLButtonElement;
    const list = requiredElement(
      requiredElement(select.parentElement).querySelector<HTMLElement>('.fc-select__list'),
    );
    const outside = document.createElement('button');
    document.body.appendChild(outside);

    button.focus();
    button.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    const option = requiredElement(
      list.querySelector<HTMLButtonElement>('.fc-select__option[data-value="d"]'),
    );
    option.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    button.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
    expect(list.hidden).toBe(false);
    outside.dispatchEvent(new MouseEvent('pointerup', { bubbles: true }));

    expect(list.hidden).toBe(true);
    outside.remove();
    handle?.dispose();
  });

  it('keeps trigger and option row button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/custom-select.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('function createCustomSelectButton');
    expect(source).toContain('function createCustomSelectOptionRow');
    expect(source).toContain('const button = createCustomSelectButton(select)');
    expect(source).toContain('const row = createCustomSelectOptionRow(opt)');
    expect(source).toContain('function optionDisabled(opt: HTMLOptionElement)');
    expect(source).toContain("opt.closest('optgroup')?.disabled === true");
    expect(source).toContain('projectDisabledState(row, optionDisabled(opt)');
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps the popup select chrome compact like Excel desktop controls', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/custom-select.css'),
      'utf8',
    );

    expect(css).toMatch(/\.fc-select__button\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(/\.fc-select__list\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(/\.fc-select__option\s*\{[\s\S]*?border-radius: 0;/);
    expect(css).toMatch(/\.fc-select__list\s*\{[\s\S]*?0 6px 16px rgba\(0, 0, 0, 0\.14\)/);
  });
});
