import { mount } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { makeHost } from './helpers.js';

const disposers: Array<{ dispose(): void }> = [];

afterEach(() => {
  for (const d of disposers.splice(0).reverse()) d.dispose();
});

/** Commands sheet protection refuses: always-refused ones and option-gated ones. */
const GATED = ['sheetBackground', 'precedents', 'insertRows', 'filter'] as const;

describe('sheet protection set through the workbook API', () => {
  it('gates commands on protect and releases them on unprotect', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    for (const id of GATED) expect(instance.commands.can(id), id).toBe(true);

    await sheet.protect();
    for (const id of GATED) expect(instance.commands.can(id), id).toBe(false);

    await sheet.unprotect();
    for (const id of GATED) expect(instance.commands.can(id), id).toBe(true);
  });

  it('releases a password-protected sheet once the password matches', async () => {
    const instance = await mount(makeHost());
    disposers.push(instance);
    const sheet = instance.workbook.sheet(0);
    await sheet.protect({ password: 'secret' });
    expect(instance.commands.can('sheetBackground')).toBe(false);
    await sheet.unprotect('secret');
    expect(instance.commands.can('sheetBackground')).toBe(true);
  });
});
