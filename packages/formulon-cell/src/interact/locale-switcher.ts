import { BUILT_IN_LOCALES, type BuiltInLocale, builtInLocaleFor } from '../i18n/locales.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';

export interface LocaleSwitcherDeps {
  /** The chrome slot reserved for this feature. */
  container: HTMLElement;
  /** Returns the locale that the mount and engine currently use. */
  getLocale: () => string;
  /** Returns true only when the mount committed the requested locale. */
  setLocale: (locale: string) => boolean;
  strings?: Strings;
  /** Receives an engine or mount exception. Rejections are surfaced as an Error. */
  onError?: (error: unknown) => void;
}

export interface LocaleSwitcherHandle {
  refresh(): void;
  setStrings(next: Strings): void;
  detach(): void;
}

const CUSTOM_OPTION_ATTRIBUTE = 'data-fc-locale-custom';
let nextLocaleSwitcherId = 0;

/**
 * Return the short option id for a built-in locale or a valid BCP 47 region
 * tag. Custom registry ids such as `de-brand` have no region and stay custom.
 */
const localeOptionFor = (locale: string): BuiltInLocale | null => {
  const mapped = builtInLocaleFor(locale);
  if (!mapped) return null;
  const normalized = locale.trim().replaceAll('_', '-');
  if (normalized.toLowerCase() === mapped) return mapped;
  try {
    const parsed = new Intl.Locale(normalized);
    const region = parsed.region;
    if (!region || parsed.script) return null;
    return parsed.baseName.toLowerCase() === `${mapped}-${region.toLowerCase()}` ? mapped : null;
  } catch {
    return null;
  }
};

const errorFor = (message: string): Error => new Error(message);

export function attachLocaleSwitcher(deps: LocaleSwitcherDeps): LocaleSwitcherHandle {
  const { container, getLocale, setLocale } = deps;
  let strings = deps.strings ?? defaultStrings;
  let disposed = false;
  let changing = false;

  container.replaceChildren();
  container.dataset.fcLocaleSwitcher = 'true';

  const label = document.createElement('label');
  label.className = 'fc-host__locale-switcher-label';
  const select = document.createElement('select');
  select.className = 'fc-host__locale-switcher-select';
  select.dataset.fcLocaleSelect = 'true';
  nextLocaleSwitcherId += 1;
  const switcherId = `fc-locale-switcher-${nextLocaleSwitcherId}`;
  select.id = `${switcherId}-select`;
  const status = document.createElement('span');
  status.className = 'fc-host__locale-switcher-status';
  status.id = `${switcherId}-status`;
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');
  status.setAttribute('aria-atomic', 'true');
  select.setAttribute('aria-describedby', status.id);
  label.htmlFor = select.id;
  container.append(label, select, status);

  const removeCustomOption = (): void => {
    select.querySelector<HTMLOptionElement>(`option[${CUSTOM_OPTION_ATTRIBUTE}]`)?.remove();
  };

  const renderOptions = (current: string): void => {
    removeCustomOption();
    select.replaceChildren();
    const selectedLocale = localeOptionFor(current);
    const optionStrings = strings.localeSwitcher.options;
    for (const locale of BUILT_IN_LOCALES) {
      const option = document.createElement('option');
      option.value = locale;
      option.lang = locale;
      option.textContent = optionStrings[locale];
      select.appendChild(option);
    }

    if (!selectedLocale) {
      const custom = document.createElement('option');
      custom.value = current;
      custom.textContent = current || '—';
      custom.setAttribute('disabled', '');
      custom.setAttribute(CUSTOM_OPTION_ATTRIBUTE, 'true');
      select.insertBefore(custom, select.firstChild);
    }

    select.value = selectedLocale ?? current;
    // A browser can reject an empty or otherwise unusual value assignment.
    // Keep the actual getter value visible through the selected option rather
    // than silently choosing one of the built-in languages.
    const expectedValue = selectedLocale ?? current;
    if (select.value !== expectedValue) {
      const fallback = Array.from(select.options).find((option) => option.value === expectedValue);
      if (fallback) fallback.selected = true;
    }
  };

  const refresh = (): void => {
    if (disposed) return;
    const current = getLocale();
    label.textContent = strings.localeSwitcher.label;
    select.setAttribute('aria-label', strings.localeSwitcher.label);
    renderOptions(current);
  };

  const clearStatus = (): void => {
    status.replaceChildren();
  };

  const showFailure = (error: unknown): void => {
    // Replace the text node on every failure so an identical retry is still
    // observable by assistive technology through the live region.
    status.replaceChildren(document.createTextNode(strings.localeSwitcher.changeFailed));
    try {
      deps.onError?.(error);
    } catch {
      // A host error hook must not interrupt restoring the selector state.
    }
  };

  const onChange = (): void => {
    if (disposed || changing) return;
    const requested = select.value;
    const previous = getLocale();
    if (!requested || requested === previous || requested === localeOptionFor(previous)) {
      clearStatus();
      refresh();
      return;
    }

    changing = true;
    clearStatus();
    try {
      let accepted = false;
      try {
        accepted = setLocale(requested);
      } catch (error) {
        refresh();
        showFailure(error);
        return;
      }

      const actual = getLocale();
      if (!accepted || localeOptionFor(actual) !== requested) {
        refresh();
        showFailure(
          errorFor(
            `${strings.localeSwitcher.changeFailed} Requested ${requested}; active locale is ${actual}.`,
          ),
        );
      } else {
        refresh();
        clearStatus();
      }
    } finally {
      changing = false;
    }
  };

  select.addEventListener('change', onChange);
  refresh();

  return {
    refresh,
    setStrings(next) {
      strings = next;
      clearStatus();
      refresh();
    },
    detach() {
      if (disposed) return;
      disposed = true;
      select.removeEventListener('change', onChange);
      container.replaceChildren();
      delete container.dataset.fcLocaleSwitcher;
    },
  };
}
