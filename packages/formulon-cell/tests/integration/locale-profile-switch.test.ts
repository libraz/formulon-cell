import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle as WorkbookHandleType } from '../../src/engine/workbook-handle.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { overlayPortalFor } from '../../src/interact/overlay-portal.js';
import {
  profileLocaleForUiLocale,
  spreadsheetProfileForLocale,
  syncLocaleProfile,
} from '../../src/mount/locale-profile.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

describe('integration: locale profile switching', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  it.each([
    ['ja', 'ja_JP'],
    ['en-US', 'en_US'],
    ['de-DE', 'de_DE'],
    ['fr-FR', 'fr_FR'],
    ['de_DE', 'de_DE'],
    ['zh-CN', 'zh_CN'],
    ['zh-TW', 'zh_TW'],
    ['zh_HK', 'zh_TW'],
    ['ko-KR', 'ko_KR'],
    ['th-TH', 'th_TH'],
  ] as const)('maps %s to the seven-locale profile suffix %s', (locale, suffix) => {
    expect(profileLocaleForUiLocale(locale)).toBe(suffix);
    expect(spreadsheetProfileForLocale(locale, 'default')).toBe(`windows-${suffix}`);
    expect(spreadsheetProfileForLocale(locale, 'mac')).toBe(`mac-${suffix}`);
  });

  it('does not override an unknown custom locale or an unsupported engine', () => {
    const recalc = vi.fn();
    const setProfile = vi.fn();
    const wb = {
      capabilities: { spreadsheetProfile: false },
      spreadsheetProfileId: vi.fn(() => null),
      setSpreadsheetProfileId: setProfile,
      recalc,
    } as unknown as WorkbookHandleType;

    expect(syncLocaleProfile(wb, 'x-private-locale', 'mac')).toBe(false);
    expect(syncLocaleProfile(wb, 'de', 'mac')).toBe(false);
    expect(setProfile).not.toHaveBeenCalled();
    expect(recalc).not.toHaveBeenCalled();
  });

  it.each(['de-brand', 'de_@@', 'de-Latn-DE', 'de-DE-posix', 'zh-Hant-TW', 'zh-Hans-CN'])(
    'does not override a custom or malformed German locale: %s',
    (locale) => {
      const recalc = vi.fn();
      const setProfile = vi.fn();
      const wb = {
        capabilities: { spreadsheetProfile: true },
        spreadsheetProfileId: vi.fn(() => 'windows-en_US'),
        setSpreadsheetProfileId: setProfile,
        recalc,
      } as unknown as WorkbookHandleType;

      expect(profileLocaleForUiLocale(locale)).toBeNull();
      expect(syncLocaleProfile(wb, locale, 'mac')).toBe(false);
      expect(setProfile).not.toHaveBeenCalled();
      expect(recalc).not.toHaveBeenCalled();
    },
  );

  it('recalculates once after a real platform profile switch', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const recalc = vi.spyOn(workbook, 'recalc');
    try {
      sheet = await mountStubSheet({ workbook, locale: 'ja', ui: { platform: 'default' } });
      expect(workbook.spreadsheetProfileId()).toBe('windows-ja_JP');
      recalc.mockClear();

      sheet.instance.setUi({ platform: 'mac' });

      expect(workbook.spreadsheetProfileId()).toBe('mac-ja_JP');
      expect(recalc).toHaveBeenCalledTimes(1);
      sheet.instance.setUi({ platform: 'mac' });
      expect(recalc).toHaveBeenCalledTimes(1);
    } finally {
      workbook.dispose();
    }
  });

  it('reselects the workbook profile when the UI locale changes', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const recalc = vi.spyOn(workbook, 'recalc');
    try {
      sheet = await mountStubSheet({ workbook, locale: 'en', ui: { platform: 'default' } });
      expect(workbook.spreadsheetProfileId()).toBe('windows-en_US');
      recalc.mockClear();

      sheet.instance.i18n.setLocale('de');

      expect(workbook.spreadsheetProfileId()).toBe('windows-de_DE');
      expect(recalc).toHaveBeenCalledTimes(1);
      sheet.instance.i18n.setLocale('de');
      expect(recalc).toHaveBeenCalledTimes(1);
    } finally {
      workbook.dispose();
    }
  });

  it('uses the mounted locale getter for direct and repeated keyboard formats', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      sheet = await mountStubSheet({ workbook, locale: 'en', ui: { platform: 'default' } });
      const dispatchKey = (event: KeyboardEventInit): void => {
        sheet?.host.dispatchEvent(new KeyboardEvent('keydown', { ...event, bubbles: true }));
      };
      const numberFormat = () => sheet?.instance.store.getState().ui.pendingFormat?.format.numFmt;

      dispatchKey({ key: '$', code: 'Digit4', ctrlKey: true, shiftKey: true });
      expect(numberFormat()).toEqual({ kind: 'currency', symbol: '$', decimals: 2 });

      expect(sheet.instance.setLocale('ja')).toBe(true);
      mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 1 });
      dispatchKey({ key: 'F4' });
      expect(numberFormat()).toEqual({ kind: 'currency', symbol: '¥', decimals: 2 });

      expect(sheet.instance.setLocale('en')).toBe(true);
      dispatchKey({ key: '#', code: 'Digit3', ctrlKey: true, shiftKey: true });
      expect(numberFormat()).toEqual({ kind: 'date', pattern: 'm/d/yyyy' });

      expect(sheet.instance.setLocale('ja')).toBe(true);
      mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 2 });
      dispatchKey({ key: 'F4' });
      expect(numberFormat()).toEqual({ kind: 'date', pattern: 'yyyy/m/d' });
    } finally {
      workbook.dispose();
    }
  });

  it('localizes stored formulas in the formula bar and inline editor', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      workbook.setFormula({ sheet: 0, row: 0, col: 0 }, '=SUM(1,2)');
      sheet = await mountStubSheet({
        workbook,
        locale: 'de',
      });
      const formulaBar = sheet.host.querySelector<HTMLTextAreaElement>(
        '.fc-host__formulabar-input',
      );
      expect(formulaBar?.value).toBe('=SUMME(1;2)');
      expect(workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1,2)');

      sheet.host.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
      const inlineEditor = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__editor');
      expect(inlineEditor?.value).toBe('=SUMME(1;2)');
    } finally {
      workbook.dispose();
    }
  });

  it.each([
    ['en', 'TRUE'],
    ['de', 'WAHR'],
    ['fr', 'VRAI'],
    ['ja', 'TRUE'],
  ] as const)('uses the %s boolean display token in the formula bar', async (locale, token) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      workbook.setBool({ sheet: 0, row: 0, col: 0 }, true);
      sheet = await mountStubSheet({
        workbook,
        locale,
      });
      const formulaBar = sheet.host.querySelector<HTMLTextAreaElement>(
        '.fc-host__formulabar-input',
      );
      expect(formulaBar?.value).toBe(token);
      expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'bool',
        value: true,
      });
    } finally {
      workbook.dispose();
    }
  });

  it.each(['de', 'fr'] as const)(
    'round-trips a decimal edit using the %s input mark',
    async (locale) => {
      const workbook = await WorkbookHandle.createDefault();
      try {
        workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1.5);
        sheet = await mountStubSheet({ workbook, locale });
        const formulaBar = sheet.host.querySelector<HTMLTextAreaElement>(
          '.fc-host__formulabar-input',
        );
        expect(formulaBar?.value).toBe('1,5');

        sheet.host.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
        const inlineEditor = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__editor');
        expect(inlineEditor?.value).toBe('1,5');
        inlineEditor?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

        expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
          kind: 'number',
          value: 1.5,
        });
      } finally {
        workbook.dispose();
      }
    },
  );

  it('exposes an atomic public locale setter', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const recalc = vi.spyOn(workbook, 'recalc');
    try {
      sheet = await mountStubSheet({ workbook, locale: 'en', ui: { platform: 'default' } });
      expect(sheet.instance.locale).toBe('en');
      recalc.mockClear();

      expect(sheet.instance.setLocale('de')).toBe(true);
      expect(sheet.instance.locale).toBe('de');
      expect(workbook.spreadsheetProfileId()).toBe('windows-de_DE');
      expect(recalc).toHaveBeenCalledTimes(1);

      expect(sheet.instance.setLocale('de')).toBe(true);
      expect(recalc).toHaveBeenCalledTimes(1);

      vi.spyOn(workbook, 'setSpreadsheetProfileId').mockReturnValue(false);
      expect(sheet.instance.setLocale('fr')).toBe(false);
      expect(sheet.instance.locale).toBe('de');
      expect(workbook.spreadsheetProfileId()).toBe('windows-de_DE');
    } finally {
      workbook.dispose();
    }
  });

  it('commits a localized formula draft before switching its profile', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      sheet = await mountStubSheet({ workbook, locale: 'de', ui: { platform: 'default' } });
      const input = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
      expect(input).not.toBeNull();
      if (!input) return;
      input.focus();
      input.dispatchEvent(new FocusEvent('focus'));
      input.value = '=SUMME(1;2)';
      input.dispatchEvent(new Event('input', { bubbles: true }));

      sheet.instance.i18n.setLocale('fr');

      expect(workbook.spreadsheetProfileId()).toBe('windows-fr_FR');
      expect(workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=SUM(1,2)');
      expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 3,
      });
    } finally {
      workbook.dispose();
    }
  });

  it('rejects a locale switch while the Mac formula palette owns a draft', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      sheet = await mountStubSheet({
        workbook,
        locale: 'en',
        ui: { platform: 'mac' },
        features: { fxDialog: true },
      });
      sheet.instance.openFunctionArguments('SUM');
      const root = sheet.host.querySelector<HTMLElement>('.fc-mac-formula-palette');
      const firstArgument = root?.querySelector<HTMLInputElement>('[data-argument-index="0"]');
      if (!root || !firstArgument) throw new Error('Mac formula palette did not open');
      firstArgument.value = '1';
      firstArgument.dispatchEvent(new Event('input', { bubbles: true }));

      expect(sheet.instance.setLocale('ja')).toBe(false);
      expect(sheet.instance.locale).toBe('en');
      expect(workbook.spreadsheetProfileId()).toBe('mac-en_US');
      expect(root.hidden).toBe(false);
      expect(firstArgument.value).toBe('1');
      expect(workbook.cellFormula({ sheet: 0, row: 0, col: 0 })).toBeNull();
    } finally {
      workbook.dispose();
    }
  });

  it.each(['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'])(
    'preserves an open modal draft in %s and switches after it closes',
    async (locale) => {
      sheet = await mountStubSheet({ locale });
      const handle = sheet.instance.features.formatDialog as unknown as {
        open(): void;
        close(): void;
      };
      handle.open();
      const overlay = overlayPortalFor(sheet.host).querySelector<HTMLElement>(
        '[class="fc-fmtdlg"]',
      );
      const bold = overlay?.querySelector<HTMLInputElement>('[data-fc-check="bold"]');
      if (!overlay || !bold) throw new Error('format dialog did not open');
      bold.checked = true;
      bold.dispatchEvent(new Event('change', { bubbles: true }));
      const next = locale === 'fr' ? 'de' : 'fr';

      expect(sheet.instance.setLocale(next)).toBe(false);
      expect(sheet.instance.locale).toBe(locale);
      // The public controller is another supported entry point.
      sheet.instance.i18n.setLocale(next);
      expect(sheet.instance.locale).toBe(locale);
      expect(overlay.isConnected).toBe(true);
      expect(overlay.hidden).toBe(false);
      expect(bold.checked).toBe(true);
      expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBeUndefined();

      handle.close();
      expect(sheet.instance.setLocale(next)).toBe(true);
      expect(sheet.instance.locale).toBe(next);
    },
  );

  it('rolls back the UI locale when a live profile write is rejected', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      sheet = await mountStubSheet({ workbook, locale: 'de', ui: { platform: 'default' } });
      vi.spyOn(workbook, 'setSpreadsheetProfileId').mockReturnValue(false);

      expect(() => sheet?.instance.i18n.setLocale('fr')).not.toThrow();
      expect(sheet.instance.i18n.locale).toBe('de');
      expect(workbook.spreadsheetProfileId()).toBe('windows-de_DE');
    } finally {
      workbook.dispose();
    }
  });

  it('rejects direct controller changes before user extension listeners observe them', async () => {
    const observed: string[] = [];
    let input: HTMLInputElement | undefined;
    sheet = await mountStubSheet({
      locale: 'en',
      extensions: [
        {
          id: 'locale-draft-owner',
          setup(ctx) {
            const dialog = document.createElement('div');
            dialog.setAttribute('role', 'dialog');
            dialog.setAttribute('aria-modal', 'true');
            input = document.createElement('input');
            input.value = 'uncommitted draft';
            dialog.append(input);
            overlayPortalFor(ctx.host).append(dialog);
            const unsubscribe = ctx.i18n.subscribe(() => {
              observed.push(ctx.i18n.locale);
              if (ctx.i18n.locale === 'fr' && input) input.value = 'lost';
            });
            return {
              dispose() {
                unsubscribe();
                dialog.remove();
              },
            };
          },
        },
      ],
    });
    sheet.instance.i18n.setLocale('fr');
    expect(sheet.instance.locale).toBe('en');
    expect(observed).not.toContain('fr');
    expect(input?.value).toBe('uncommitted draft');
  });

  it('coalesces same-language dictionary updates until the final draft closes', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const handle = sheet.instance.features.formatDialog as unknown as {
      open(): void;
      close(): void;
    };
    handle.open();
    const portal = overlayPortalFor(sheet.host);
    const overlay = portal.querySelector<HTMLElement>('[class="fc-fmtdlg"]');
    const bold = overlay?.querySelector<HTMLInputElement>('[data-fc-check="bold"]');
    if (!overlay || !bold) throw new Error('format dialog did not open');
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    const secondDraft = document.createElement('div');
    secondDraft.className = 'fc-cmtnote';
    portal.append(secondDraft);
    sheet.instance.i18n.extend('en', {
      a11y: { spreadsheet: 'First label' },
      formatDialog: { title: 'First format title' },
    });
    sheet.instance.i18n.extend('en', {
      a11y: { spreadsheet: 'Latest label' },
      formatDialog: { title: 'Latest format title' },
    });
    expect(sheet.host.getAttribute('aria-label')).toBe('Latest label');
    expect(overlay.isConnected).toBe(true);
    expect(bold.checked).toBe(true);
    handle.close();
    await Promise.resolve();
    expect(overlay.isConnected).toBe(true);
    secondDraft.remove();
    await vi.waitFor(() => expect(overlay.isConnected).toBe(false));
    const current = sheet.instance.features.formatDialog as unknown as {
      open(): void;
      close(): void;
    };
    current.open();
    expect(portal.querySelector('[class="fc-fmtdlg"]')?.getAttribute('aria-label')).toBe(
      'Latest format title',
    );
    current.close();
  });

  it('flushes deferred translations after a draft closes through a CSS class', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const portal = overlayPortalFor(sheet.host);
    const style = document.createElement('style');
    style.textContent = '.fc-test-closed-draft { display: none; }';
    document.head.append(style);
    const draft = document.createElement('div');
    draft.className = 'fc-cmtnote';
    portal.append(draft);
    try {
      await Promise.resolve();
      const changed = vi.fn();
      const unsubscribe = sheet.instance.on('localeChange', changed);
      sheet.instance.i18n.extend('en', { formatDialog: { title: 'Updated format title' } });
      expect(changed).not.toHaveBeenCalled();
      draft.classList.add('fc-test-closed-draft');
      await vi.waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
      const current = sheet.instance.features.formatDialog as unknown as {
        open(): void;
        close(): void;
      };
      current.open();
      expect(portal.querySelector('[class="fc-fmtdlg"]')?.getAttribute('aria-label')).toBe(
        'Updated format title',
      );
      current.close();
      unsubscribe();
    } finally {
      draft.remove();
      style.remove();
    }
  });

  it('restores the prior profile after an unexpected readback', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      const setProfile = vi.spyOn(workbook, 'setSpreadsheetProfileId');
      const readProfile = vi
        .spyOn(workbook, 'spreadsheetProfileId')
        .mockImplementationOnce(() => 'windows-ja_JP')
        .mockImplementationOnce(() => 'windows-en_US');
      const recalc = vi.spyOn(workbook, 'recalc');

      expect(syncLocaleProfile(workbook, 'de', 'default')).toBe(false);
      expect(setProfile.mock.calls.map(([profile]) => profile)).toEqual([
        'windows-de_DE',
        'windows-ja_JP',
      ]);
      expect(recalc).toHaveBeenCalledTimes(1);
      expect(readProfile).toHaveBeenCalled();
      expect(workbook.spreadsheetProfileId()).toBe('windows-ja_JP');
    } finally {
      workbook.dispose();
    }
  });

  it('restores the prior profile when the target recalc fails', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
      const setProfile = vi.spyOn(workbook, 'setSpreadsheetProfileId');
      let recalcCalls = 0;
      vi.spyOn(workbook, 'recalc').mockImplementation(() => {
        recalcCalls += 1;
        if (recalcCalls === 1) throw new Error('target recalc failed');
      });

      expect(() => syncLocaleProfile(workbook, 'de', 'default', { throwOnFailure: true })).toThrow(
        /target recalc failed/,
      );
      expect(setProfile.mock.calls.map(([profile]) => profile)).toEqual([
        'windows-de_DE',
        'windows-ja_JP',
      ]);
      expect(recalcCalls).toBe(2);
      expect(workbook.spreadsheetProfileId()).toBe('windows-ja_JP');
    } finally {
      workbook.dispose();
    }
  });
});
