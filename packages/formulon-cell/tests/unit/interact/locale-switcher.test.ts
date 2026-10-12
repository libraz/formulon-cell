import { afterEach, describe, expect, it, vi } from 'vitest';
import { de, en, ja } from '../../../src/i18n/strings.js';
import { attachLocaleSwitcher } from '../../../src/interact/locale-switcher.js';

describe('attachLocaleSwitcher', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('renders all seven autonym options and changes through the host callback', () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    let locale = 'ja';
    const setLocale = vi.fn((next: string) => {
      locale = next;
      return true;
    });
    const handle = attachLocaleSwitcher({
      container,
      getLocale: () => locale,
      setLocale,
      strings: ja,
    });

    const select = container.querySelector<HTMLSelectElement>('select');
    if (!select) throw new Error('locale select was not rendered');
    expect([...select.options].map((option) => option.value)).toEqual([
      'en',
      'ja',
      'de',
      'fr',
      'zh',
      'ko',
      'th',
    ]);
    expect([...select.options].map((option) => option.lang)).toEqual([
      'en',
      'ja',
      'de',
      'fr',
      'zh',
      'ko',
      'th',
    ]);
    expect(select?.value).toBe('ja');
    expect(select?.getAttribute('aria-label')).toBe('言語');
    expect([...select.options].map((option) => option.textContent)).toContain('Deutsch');

    select.value = 'de';
    select.dispatchEvent(new Event('change', { bubbles: true }));
    expect(setLocale).toHaveBeenCalledWith('de');
    expect(locale).toBe('de');
    expect(select?.value).toBe('de');
    expect(container.querySelector('.fc-host__locale-switcher-status')?.textContent).toBe('');

    handle.detach();
  });

  it('restores the active locale and announces a rejected change', () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onError = vi.fn();
    const locale = 'ja';
    const selectLocale = attachLocaleSwitcher({
      container,
      getLocale: () => locale,
      setLocale: () => false,
      strings: en,
      onError,
    });
    const select = container.querySelector<HTMLSelectElement>('select');
    if (!select) throw new Error('locale select was not rendered');

    select.value = 'fr';
    select.dispatchEvent(new Event('change'));

    expect(select.value).toBe('ja');
    expect(container.querySelector('.fc-host__locale-switcher-status')?.textContent).toBe(
      en.localeSwitcher.changeFailed,
    );
    expect(onError).toHaveBeenCalledOnce();
    selectLocale.detach();
  });

  it('keeps a custom active locale as a disabled transient option', () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const handle = attachLocaleSwitcher({
      container,
      getLocale: () => 'de-brand',
      setLocale: () => true,
      strings: en,
    });
    const custom = container.querySelector<HTMLOptionElement>('[data-fc-locale-custom]');

    expect(custom?.value).toBe('de-brand');
    expect(custom?.disabled).toBe(true);
    expect(container.querySelector<HTMLSelectElement>('select')?.value).toBe('de-brand');
    expect(container.querySelectorAll('option')).toHaveLength(8);

    handle.setStrings(de);
    expect(container.querySelector<HTMLSelectElement>('select')?.getAttribute('aria-label')).toBe(
      'Sprache',
    );
    expect(container.querySelector<HTMLOptionElement>('[data-fc-locale-custom]')?.disabled).toBe(
      true,
    );
    handle.detach();
  });

  it.each([
    ['en-US', 'en'],
    ['de_DE', 'de'],
  ])('maps supported regional locale %s to the %s option', (current, selected) => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const handle = attachLocaleSwitcher({
      container,
      getLocale: () => current,
      setLocale: () => true,
      strings: en,
    });
    const select = container.querySelector<HTMLSelectElement>('select');
    if (!select) throw new Error('locale select was not rendered');

    expect(select.options).toHaveLength(7);
    expect(select.value).toBe(selected);
    expect(container.querySelector('[data-fc-locale-custom]')).toBeNull();
    handle.detach();
  });

  it('restores the getter value when the host callback throws', () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const error = new Error('profile rejected');
    const onError = vi.fn();
    const handle = attachLocaleSwitcher({
      container,
      getLocale: () => 'ja',
      setLocale: () => {
        throw error;
      },
      strings: en,
      onError,
    });
    const select = container.querySelector<HTMLSelectElement>('select');
    if (!select) throw new Error('locale select was not rendered');

    select.value = 'en';
    expect(() => select.dispatchEvent(new Event('change'))).not.toThrow();
    expect(select.value).toBe('ja');
    expect(onError).toHaveBeenCalledWith(error);
    expect(container.querySelector('.fc-host__locale-switcher-status')?.textContent).toBe(
      en.localeSwitcher.changeFailed,
    );
    handle.detach();
  });

  it('links a stable status live region and refreshes it for repeated failures', () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    let locale = 'ja';
    let accept = false;
    const handle = attachLocaleSwitcher({
      container,
      getLocale: () => locale,
      setLocale: (next) => {
        if (!accept) return false;
        locale = next;
        return true;
      },
      strings: en,
    });
    const select = container.querySelector<HTMLSelectElement>('select');
    const status = container.querySelector<HTMLElement>('.fc-host__locale-switcher-status');
    if (!select || !status) throw new Error('locale switcher controls were not rendered');

    expect(status.id).toMatch(/^fc-locale-switcher-\d+-status$/);
    expect(select.getAttribute('aria-describedby')).toBe(status.id);
    expect(select.hasAttribute('aria-invalid')).toBe(false);

    select.value = 'fr';
    select.dispatchEvent(new Event('change'));
    const firstFailure = status.firstChild;
    expect(firstFailure).not.toBeNull();
    expect(status.textContent).toBe(en.localeSwitcher.changeFailed);

    select.value = 'fr';
    select.dispatchEvent(new Event('change'));
    expect(status.firstChild).not.toBe(firstFailure);
    expect(status.textContent).toBe(en.localeSwitcher.changeFailed);

    select.value = 'ja';
    select.dispatchEvent(new Event('change'));
    expect(status.textContent).toBe('');

    select.value = 'fr';
    accept = true;
    select.dispatchEvent(new Event('change'));
    expect(locale).toBe('fr');
    expect(status.textContent).toBe('');

    select.focus();
    handle.setStrings(de);
    expect(document.activeElement).toBe(select);
    expect(status.textContent).toBe('');
    handle.detach();
  });

  it('uses unique label targets for separate spreadsheet instances', () => {
    const firstContainer = document.createElement('div');
    const secondContainer = document.createElement('div');
    document.body.append(firstContainer, secondContainer);
    const first = attachLocaleSwitcher({
      container: firstContainer,
      getLocale: () => 'ja',
      setLocale: () => true,
    });
    const second = attachLocaleSwitcher({
      container: secondContainer,
      getLocale: () => 'en',
      setLocale: () => true,
    });
    const firstSelect = firstContainer.querySelector<HTMLSelectElement>('select');
    const secondSelect = secondContainer.querySelector<HTMLSelectElement>('select');
    const firstLabel = firstContainer.querySelector('label');
    const secondLabel = secondContainer.querySelector('label');
    if (!firstSelect || !secondSelect || !firstLabel || !secondLabel) {
      throw new Error('locale switcher controls were not rendered');
    }

    expect(firstSelect.id).not.toBe(secondSelect.id);
    expect(firstLabel.htmlFor).toBe(firstSelect.id);
    expect(secondLabel.htmlFor).toBe(secondSelect.id);
    expect(firstSelect.getAttribute('aria-describedby')).toBe(
      firstContainer.querySelector<HTMLElement>('.fc-host__locale-switcher-status')?.id,
    );
    expect(secondSelect.getAttribute('aria-describedby')).toBe(
      secondContainer.querySelector<HTMLElement>('.fc-host__locale-switcher-status')?.id,
    );
    first.detach();
    second.detach();
  });

  it('removes the event listener and contents on detach', () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const setLocale = vi.fn(() => true);
    const handle = attachLocaleSwitcher({
      container,
      getLocale: () => 'ja',
      setLocale,
    });
    const select = container.querySelector<HTMLSelectElement>('select');
    if (!select) throw new Error('locale select was not rendered');

    handle.detach();
    select.value = 'en';
    select.dispatchEvent(new Event('change'));

    expect(setLocale).not.toHaveBeenCalled();
    expect(container.childElementCount).toBe(0);
  });
});
