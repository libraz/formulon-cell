import '../../api/setup-dom.js';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { type Workbook, workbookOf } from '../../../src/api/workbook.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { Extension, ExtensionContext } from '../../../src/extensions/types.js';
import { mount } from '../../../src/index.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';

describe('extension public context', () => {
  beforeEach(() => installDomStubs());
  afterEach(() => uninstallDomStubs());

  it('hands user extensions the facade workbook and a read-only i18n view', async () => {
    const { host, cleanup } = createHostElement();
    const first = workbookOf(await WorkbookHandle.createDefault({ preferStub: true }));
    const second = workbookOf(await WorkbookHandle.createDefault({ preferStub: true }));
    let ctx: ExtensionContext | undefined;
    const swaps: Workbook[] = [];
    const ext: Extension = {
      id: 'probe',
      setup(c) {
        ctx = c;
        return { dispose: c.spreadsheet.on('workbook', (e) => swaps.push(e.workbook)) };
      },
    };

    const instance = await mount(host, { workbook: first, locale: 'en', extensions: [ext] });
    try {
      if (!ctx) throw new Error('extension was not set up');
      expect(ctx.workbook).toBe(instance.workbook);
      expect(ctx.i18n.locale).toBe(instance.i18n.locale);
      expect(ctx.i18n).not.toHaveProperty('extend');
      expect(ctx.i18n).not.toHaveProperty('register');
      expect(ctx.i18n).not.toHaveProperty('setLocale');

      const swapped = new Promise((resolve) => instance.on('workbook', resolve));
      instance.update({ workbook: second });
      await swapped;

      expect(swaps).toEqual([second]);
      expect(ctx.workbook).toBe(second);
      expect(ctx.workbook).toBe(instance.workbook);
    } finally {
      instance.dispose();
      cleanup();
    }
  });
});
