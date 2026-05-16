// Shell locale + autosave switch helpers.
// Extracted from main.ts. The factory pattern lets the host pass in live
// references (autosave button, localized shell text, document root, autosave
// state accessor) without coupling this module to global state.

export interface ShellLocaleCtx<Key extends string> {
  autosaveSwitch: HTMLButtonElement | null;
  shellText: Record<Key, string> & {
    autosave: string;
    autosaveOn: string;
    autosaveOff: string;
  };
  html: HTMLElement;
  ribbonLang: 'ja' | 'en';
  getAutosaveEnabled: () => boolean;
}

export interface ShellLocaleApi {
  setShellLabel: (element: Element, value: string) => void;
  refreshAutosave: () => void;
  applyShellLocale: () => void;
}

export const createShellLocale = <Key extends string>(ctx: ShellLocaleCtx<Key>): ShellLocaleApi => {
  const { autosaveSwitch, shellText, html, ribbonLang, getAutosaveEnabled } = ctx;

  const setShellLabel = (element: Element, value: string): void => {
    element.setAttribute('aria-label', value);
    if (element instanceof HTMLElement) element.title = value;
  };

  const refreshAutosave = (): void => {
    if (!autosaveSwitch) return;
    const autosaveEnabled = getAutosaveEnabled();
    autosaveSwitch.setAttribute('aria-pressed', autosaveEnabled ? 'true' : 'false');
    autosaveSwitch.classList.toggle('app__autosave-switch--on', autosaveEnabled);
    autosaveSwitch.title = autosaveEnabled ? shellText.autosaveOn : shellText.autosaveOff;
    autosaveSwitch.setAttribute(
      'aria-label',
      `${shellText.autosave}: ${autosaveEnabled ? shellText.autosaveOn : shellText.autosaveOff}`,
    );
  };

  const applyShellLocale = (): void => {
    html.lang = ribbonLang === 'ja' ? 'ja' : 'en';
    for (const el of document.querySelectorAll<HTMLElement>('[data-shell-i18n]')) {
      const key = el.dataset.shellI18n as Key | undefined;
      if (key && shellText[key]) el.textContent = shellText[key];
    }
    for (const el of document.querySelectorAll<HTMLElement>('[data-shell-i18n-label]')) {
      const key = el.dataset.shellI18nLabel as Key | undefined;
      if (key && shellText[key]) setShellLabel(el, shellText[key]);
    }
    for (const el of document.querySelectorAll<HTMLInputElement>('[data-shell-i18n-placeholder]')) {
      const key = el.dataset.shellI18nPlaceholder as Key | undefined;
      if (key && shellText[key]) el.placeholder = shellText[key];
    }
    for (const el of document.querySelectorAll<HTMLElement>('[data-shell-i18n-title]')) {
      const key = el.dataset.shellI18nTitle as Key | undefined;
      if (key && shellText[key]) el.title = shellText[key];
    }
    refreshAutosave();
  };

  return { setShellLabel, refreshAutosave, applyShellLocale };
};
