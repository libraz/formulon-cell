/** Consumer tests for `saveWithPrompts`, the Save As flow of the `./toolbar` entry. */
import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { saveWithPrompts } from '@libraz/formulon-cell/toolbar';
import { afterEach, describe, expect, it } from 'vitest';
import { makeHost } from './helpers.js';

const live: SpreadsheetInstance[] = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en' });
  live.push(instance);
  return instance;
};

const flush = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

const clickButton = (within: Element, label: string): void => {
  const button = [...within.querySelectorAll<HTMLButtonElement>('button')].find(
    (b) => b.textContent === label,
  );
  if (!button) throw new Error(`no button ${label}`);
  button.click();
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
  for (const el of document.querySelectorAll('[role="alertdialog"], [role="dialog"]')) el.remove();
});

describe('saveWithPrompts', () => {
  it('writes the given format without a prompt when nothing is lost', async () => {
    const sheet = await mountSheet();
    sheet.workbook.sheet(0).setNumber('A1', 42);
    const outcome = await saveWithPrompts(sheet, { format: 'xlsx' });
    expect(outcome?.format).toBe('xlsx');
    expect(outcome?.bytes.length).toBeGreaterThan(0);
    expect(document.querySelector('[role="alertdialog"]')).toBeNull();
  });

  it('asks for the file type itself when format is omitted', async () => {
    const sheet = await mountSheet();
    const pending = saveWithPrompts(sheet);
    await flush();
    const select = document.querySelector<HTMLSelectElement>('[data-fc-save-format]');
    expect(select).not.toBeNull();
    clickButton(select?.closest('[role="dialog"]') ?? document.body, 'Save');
    expect((await pending)?.format).toBe('xlsx');
  });

  it('resolves null when the file type choice is cancelled', async () => {
    const sheet = await mountSheet();
    const pending = saveWithPrompts(sheet, { host: sheet.host });
    await flush();
    const select = document.querySelector<HTMLSelectElement>('[data-fc-save-format]');
    clickButton(select?.closest('[role="dialog"]') ?? document.body, 'Cancel');
    expect(await pending).toBeNull();
  });
});
