import { afterEach, describe, expect, it } from 'vitest';
import { gotoSpecial } from '../../../src/extensions/built-ins.js';
import type { BuiltInExtension, BuiltInExtensionHandle } from '../../../src/extensions/types.js';
import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

let sheet: MountedStubSheet | undefined;

afterEach(() => {
  sheet?.dispose();
  sheet = undefined;
  document.body.replaceChildren();
});

describe('dialog features and the UI language', () => {
  it('open in the strings handed to setStrings', async () => {
    sheet = await mountStubSheet({ locale: 'en', features: { gotoSpecial: false } });
    const { instance, host } = sheet;
    const ext = gotoSpecial() as unknown as BuiltInExtension;
    const handle = ext.setup({
      host,
      store: instance.store,
      history: instance.history,
      i18n: instance.i18n,
      workbook: instance.workbook,
    } as unknown as Parameters<BuiltInExtension['setup']>[0]) as BuiltInExtensionHandle;
    const label = () => document.querySelector('.fc-goto')?.getAttribute('aria-label');
    const dialog = handle as unknown as { open(): void; close(): void };
    try {
      dialog.open();
      expect(label()).toBe(en.goToDialog.title);
      dialog.close();
      handle.setStrings?.(ja);
      dialog.open();
      expect(label()).toBe(ja.goToDialog.title);
    } finally {
      handle.dispose();
    }
  });
});
