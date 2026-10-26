import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { setActive } from './fixtures.js';

const A1 = addrKey({ sheet: 0, row: 0, col: 0 });

const GRADIENT: NonNullable<CellFormat['fillGradient']> = {
  type: 'linear',
  degree: 90,
  left: 0,
  right: 0,
  top: 0,
  bottom: 0,
  stops: [
    { position: 0, color: '#ff0000' },
    { position: 1, color: '#0000ff' },
  ],
};

const q = <T extends Element>(selector: string): T => {
  const el = document.querySelector<T>(selector);
  if (!el) throw new Error(`missing ${selector}`);
  return el;
};

const click = (el: Element): void => {
  el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
};

const check = (selector: string, checked: boolean): void => {
  const input = q<HTMLInputElement>(selector);
  input.checked = checked;
  input.dispatchEvent(new Event('change', { bubbles: true }));
};

const setColor = (selector: string, value: string): void => {
  const input = q<HTMLInputElement>(selector);
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
};

const ok = (): void => click(q('.fc-fmtdlg__btn--primary'));

describe('Format Cells font effects and Fill Effects', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  const seed = (format: CellFormat): void => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set(A1, format);
      return { ...s, format: { ...s.format, formats } };
    });
  };
  const formatA1 = (): CellFormat | undefined => store.getState().format.formats.get(A1);

  beforeEach(() => {
    host = document.createElement('div');
    host.className = 'fc-host';
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setActive(store, 0, 0);
  });

  afterEach(() => {
    document.body.replaceChildren();
  });

  it('offers Outline and Shadow on the Mac and writes them', () => {
    host.dataset.fcPlatform = 'mac';
    const handle = attachFormatDialog({ host, store });
    handle.open('font');
    expect(q<HTMLInputElement>('input[data-fc-check="fontOutline"]').closest('label')?.hidden).toBe(
      false,
    );
    check('input[data-fc-check="fontOutline"]', true);
    check('input[data-fc-check="fontShadow"]', true);
    ok();
    expect(formatA1()).toMatchObject({ fontOutline: true, fontShadow: true });
    handle.detach();
  });

  it('hides Outline and Shadow off the Mac', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('font');
    expect(q<HTMLInputElement>('input[data-fc-check="fontOutline"]').closest('label')?.hidden).toBe(
      true,
    );
    expect(q<HTMLInputElement>('input[data-fc-check="fontShadow"]').closest('label')?.hidden).toBe(
      true,
    );
    handle.detach();
  });

  it('keeps loaded font effects and a gradient when another property changes', () => {
    seed({ fontOutline: true, fontCondense: true, fillGradient: GRADIENT });
    const handle = attachFormatDialog({ host, store });
    handle.open('font');
    check('input[data-fc-check="bold"]', true);
    ok();
    expect(formatA1()).toMatchObject({
      bold: true,
      fontOutline: true,
      fontCondense: true,
      fillGradient: GRADIENT,
    });
    handle.detach();
  });

  it('authors a two-colour vertical gradient through Fill Effects', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('fill');
    click(q('button[data-fc-fill-effects]'));
    setColor('input[data-fc-gradient-color="1"]', '#ffffff');
    setColor('input[data-fc-gradient-color="2"]', '#4472c4');
    check('input[data-fc-gradient-style="vertical"]', true);
    click(q('button[data-fc-gradient-variant="0"]'));
    ok();
    expect(formatA1()?.fillGradient).toEqual({
      type: 'linear',
      degree: 0,
      left: 0,
      right: 0,
      top: 0,
      bottom: 0,
      stops: [
        { position: 0, color: '#ffffff' },
        { position: 1, color: '#4472c4' },
      ],
    });
    expect(formatA1()?.fill).toBeUndefined();
    handle.detach();
  });

  it('authors a gradient from the centre', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('fill');
    click(q('button[data-fc-fill-effects]'));
    setColor('input[data-fc-gradient-color="1"]', '#ffffff');
    setColor('input[data-fc-gradient-color="2"]', '#000000');
    check('input[data-fc-gradient-style="fromCenter"]', true);
    click(q('button[data-fc-gradient-variant="1"]'));
    ok();
    expect(formatA1()?.fillGradient).toMatchObject({
      type: 'path',
      left: 0.5,
      right: 0.5,
      top: 0.5,
      bottom: 0.5,
      stops: [
        { position: 0, color: '#000000' },
        { position: 1, color: '#ffffff' },
      ],
    });
    handle.detach();
  });

  it('opens Fill Effects on a loaded gradient and a solid colour replaces it', () => {
    seed({ fillGradient: GRADIENT });
    const handle = attachFormatDialog({ host, store });
    handle.open('fill');
    expect(q<HTMLElement>('.fc-fmtdlg__fill-effects').hidden).toBe(false);
    expect(q<HTMLInputElement>('input[data-fc-gradient-color="1"]').value).toBe('#ff0000');
    setColor('input[data-fc-color="fill"]', '#00ff00');
    ok();
    expect(formatA1()?.fill).toBe('#00ff00');
    expect(formatA1()?.fillGradient).toBeUndefined();
    handle.detach();
  });

  it('conditional-format formats offer neither Fill Effects nor Outline/Shadow and keep a solid fill', () => {
    host.dataset.fcPlatform = 'mac';
    const handle = attachFormatDialog({ host, store });
    let applied: Partial<CellFormat> | null = null;
    handle.open('fill', {
      mode: 'dxf',
      initialFormat: { fill: '#00ff00' },
      onApplyDxf: (format) => {
        applied = format;
      },
    });
    expect(q<HTMLButtonElement>('button[data-fc-fill-effects]').hidden).toBe(true);
    expect(q<HTMLElement>('.fc-fmtdlg__fill-effects').hidden).toBe(true);
    expect(q<HTMLInputElement>('input[data-fc-check="fontOutline"]').closest('label')?.hidden).toBe(
      true,
    );
    expect(q<HTMLInputElement>('input[data-fc-check="fontShadow"]').closest('label')?.hidden).toBe(
      true,
    );
    ok();
    expect(applied).toMatchObject({ fill: '#00ff00' });
    handle.detach();
  });

  it('a regular Format Cells open after a conditional-format one offers Fill Effects again', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('fill', { mode: 'dxf', onApplyDxf: () => undefined });
    handle.close();
    handle.open('fill');
    expect(q<HTMLButtonElement>('button[data-fc-fill-effects]').hidden).toBe(false);
    handle.detach();
  });
});
