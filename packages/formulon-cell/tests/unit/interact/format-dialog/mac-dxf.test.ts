import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { ja } from '../../../../src/i18n/strings.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import type { MacDxfForm } from '../../../../src/interact/format-dialog-tabs/controller.js';
import type { CellFormat } from '../../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { setActive } from './fixtures.js';

const mount = (platform: 'mac' | 'default') => {
  const fcHost = document.createElement('div');
  fcHost.className = 'fc-host';
  fcHost.dataset.fcPlatform = platform;
  const host = document.createElement('div');
  host.tabIndex = -1;
  fcHost.appendChild(host);
  document.body.appendChild(fcHost);
  const store = createSpreadsheetStore();
  setActive(store, 0, 0);
  return attachFormatDialog({ host, store, strings: ja });
};

const q = <T extends HTMLElement>(selector: string): T => {
  const el = document.querySelector<T>(selector);
  if (!el) throw new Error(`missing ${selector}`);
  return el;
};

const visibleTabs = (): string[] =>
  [...document.querySelectorAll<HTMLElement>('.fc-fmtdlg__tabs [role="tab"]')]
    .filter((b) => !b.hidden)
    .map((b) => b.textContent ?? '');

const openDxf = (
  platform: 'mac' | 'default',
  initialFormat: Partial<CellFormat> = {},
  dxfForm: MacDxfForm = 'styleElement',
  onApplyDxf: (f: Partial<CellFormat>) => void = () => undefined,
) => {
  const handle = mount(platform);
  // Conditional formatting asks for its rule's format on the number tab.
  handle.open(dxfForm === 'styleElement' ? 'font' : 'number', {
    mode: 'dxf',
    dxfForm,
    initialFormat,
    onApplyDxf,
  });
  return handle;
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('table style element Format dialog on the Mac profile', () => {
  it('offers exactly the font, border and fill tabs under the dialog title', () => {
    openDxf('mac');
    expect(q('.fc-fmtdlg__title').textContent).toBe('セルの書式設定');
    expect(visibleTabs()).toEqual(['フォント', '罫線', '塗りつぶし']);
  });

  it('opens on the font tab even when no tab is named', () => {
    const handle = mount('mac');
    handle.open(undefined, {
      mode: 'dxf',
      dxfForm: 'styleElement',
      initialFormat: {},
      onApplyDxf: () => undefined,
    });
    expect(q('[role="tab"][aria-selected="true"]').textContent).toBe('フォント');
  });

  it('greys the font name and size lists and leaves only the style list active', () => {
    openDxf('mac');
    expect(q<HTMLInputElement>('input[data-fc-input="family"]').disabled).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-input="size"]').disabled).toBe(true);
    for (const item of document.querySelectorAll<HTMLButtonElement>(
      '.fc-fmtdlg__font-list--family button, .fc-fmtdlg__font-list--size button',
    )) {
      expect(item.disabled).toBe(true);
    }
    for (const item of document.querySelectorAll<HTMLButtonElement>(
      '.fc-fmtdlg__font-list--style button',
    )) {
      expect(item.disabled).toBe(false);
    }
  });

  it('shows strikethrough mixed while unset and disables superscript and subscript', () => {
    openDxf('mac');
    expect(q<HTMLInputElement>('input[data-fc-check="strike"]').indeterminate).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="superscript"]').disabled).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="subscript"]').disabled).toBe(true);
  });

  it('shows strikethrough definite once the format sets it', () => {
    openDxf('mac', { strike: true });
    const strike = q<HTMLInputElement>('input[data-fc-check="strike"]');
    expect(strike.indeterminate).toBe(false);
    expect(strike.checked).toBe(true);
  });

  it('keeps an untouched strikethrough out of the applied format', () => {
    let applied: Partial<CellFormat> | undefined;
    openDxf('mac', { bold: true }, 'styleElement', (f) => {
      applied = f;
    });
    q('.fc-fmtdlg__btn--primary').dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(applied?.bold).toBe(true);
    expect(applied?.strike).toBeUndefined();
  });

  it('uses a popup for underline and automatic for the color', () => {
    openDxf('mac');
    expect(q<HTMLSelectElement>('select[data-fc-input="underline"]').disabled).toBe(false);
    expect(q('input[data-fc-color="font"]').dataset.fcAutomatic).toBe('true');
    document.body.innerHTML = '';
    openDxf('mac', { color: '#ff0000' });
    expect(q('input[data-fc-color="font"]').dataset.fcAutomatic).toBeUndefined();
  });

  it('has a clear button that is disabled until something is set, then resets', () => {
    openDxf('mac');
    const clear = q<HTMLButtonElement>('.fc-fmtdlg__btn--clear');
    expect(clear.hidden).toBe(false);
    expect(clear.textContent).toBe('クリア');
    expect(clear.disabled).toBe(true);
    const bold = q<HTMLInputElement>('input[data-fc-check="bold"]');
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    expect(clear.disabled).toBe(false);
    clear.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(bold.checked).toBe(false);
    expect(clear.disabled).toBe(true);
    expect(q('[role="tab"][aria-selected="true"]').textContent).toBe('フォント');
  });

  it('starts with the clear button enabled when the format already sets something', () => {
    openDxf('mac', { italic: true });
    expect(q<HTMLButtonElement>('.fc-fmtdlg__btn--clear').disabled).toBe(false);
  });
});

// Excel for Mac, New Rule > Classic > Custom Format...: the element form's locks, with Number kept
// and the underline fixed as well.
describe('conditional format Custom Format dialog on the Mac profile', () => {
  it('keeps the number tab and opens on the font tab', () => {
    openDxf('mac', {}, 'conditional');
    expect(visibleTabs()).toEqual(['表示形式', 'フォント', '罫線', '塗りつぶし']);
    expect(q('[role="tab"][aria-selected="true"]').textContent).toBe('フォント');
  });

  it('greys the font name, size, underline and super/subscript, with strikethrough mixed', () => {
    openDxf('mac', {}, 'conditional');
    expect(q<HTMLInputElement>('input[data-fc-input="family"]').disabled).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-input="size"]').disabled).toBe(true);
    expect(q<HTMLSelectElement>('select[data-fc-input="underline"]').disabled).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="superscript"]').disabled).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="subscript"]').disabled).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="strike"]').indeterminate).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="bold"]').disabled).toBe(false);
    expect(q('input[data-fc-color="font"]').dataset.fcAutomatic).toBe('true');
  });

  it('greys Clear until a change', () => {
    openDxf('mac', {}, 'conditional');
    const clear = q<HTMLButtonElement>('.fc-fmtdlg__btn--clear');
    expect(clear.hidden).toBe(false);
    expect(clear.disabled).toBe(true);
    const italic = q<HTMLInputElement>('input[data-fc-check="italic"]');
    italic.checked = true;
    italic.dispatchEvent(new Event('change', { bubbles: true }));
    expect(clear.disabled).toBe(false);
  });
});

describe('Mac differential format layout', () => {
  const css = readFileSync(
    join(
      resolve(dirname(fileURLToPath(import.meta.url)), '../../../..'),
      'src/styles/core/app/format-dialog/mac-layout.css',
    ),
    'utf8',
  );
  const leftOf = (selectorEnd: string): string | undefined =>
    [...css.matchAll(/([^{}]+)\{([^}]*)\}/g)]
      .filter(([, selector]) => selector?.replace(/\s+/g, ' ').trim().endsWith(selectorEnd))
      .map(([, , body]) => body?.match(/(?:^|[;\s])left:\s*([^;]+);/)?.[1]?.trim())
      .filter((left) => left !== undefined)
      .at(-1);

  // Excel for Mac: Clear sits at the tab panel's bottom-right corner, the footer holding only
  // Cancel and OK.
  it('places Clear in the tab panel, above the footer and off Cancel', () => {
    const cancel = leftOf('.fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)');
    const clear = leftOf('.fc-fmtdlg__btn.fc-fmtdlg__btn--clear');
    expect(cancel).toBeDefined();
    expect(clear).toBe('501px');
    const clearRule = [...css.matchAll(/([^{}]+)\{([^}]*)\}/g)].find(([, selector]) =>
      selector?.replace(/\s+/g, ' ').trim().endsWith('.fc-fmtdlg__btn.fc-fmtdlg__btn--clear'),
    );
    // The tab panel ends at 553 (top 38, height 515); the footer row starts at 558.
    const top = Number.parseFloat(clearRule?.[2]?.match(/top:\s*([\d.]+)px/)?.[1] ?? 'NaN');
    expect(top + 26).toBeLessThanOrEqual(553);
  });

  it('marks the overlay of a Mac differential format, so the Number hint leaves room for Clear', () => {
    openDxf('mac', {}, 'conditional');
    expect(document.querySelector('.fc-fmtdlg--mac-dxf')).not.toBeNull();
    expect(css).toMatch(
      /\.fc-fmtdlg--mac\.fc-fmtdlg--mac-dxf[\s\S]*?\.fc-fmtdlg__hintbar[^{]*\{\s*width: 464px;/,
    );
  });

  it('puts the Automatic caption in the color popup, off the Color label', () => {
    const label = leftOf('.fc-fmtdlg__font-color-row > span');
    const caption = leftOf('.fc-fmtdlg__font-color-row > .fc-fmtdlg__color-auto');
    expect(caption).toBeDefined();
    expect(caption).not.toBe(label);
  });
});

describe('differential format dialog elsewhere', () => {
  it('keeps the full form for a conditional-format rule on the default profile', () => {
    openDxf('default', {}, 'conditional');
    expect(q('[role="tab"][aria-selected="true"]').textContent).toBe('表示形式');
    expect(q<HTMLButtonElement>('.fc-fmtdlg__btn--clear').hidden).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-input="family"]').disabled).toBe(false);
    expect(q<HTMLSelectElement>('select[data-fc-input="underline"]').disabled).toBe(false);
  });

  it('is unchanged on the default profile', () => {
    openDxf('default');
    expect(visibleTabs()).toContain('表示形式');
    expect(q<HTMLInputElement>('input[data-fc-input="family"]').disabled).toBe(false);
    expect(q<HTMLInputElement>('input[data-fc-check="strike"]').indeterminate).toBe(false);
    expect(q<HTMLButtonElement>('.fc-fmtdlg__btn--clear').hidden).toBe(true);
  });
});
