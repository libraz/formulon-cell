import { afterEach, describe, expect, it, vi } from 'vitest';
import { createI18nController } from '../../../src/i18n/controller.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { cancelOpenAppDialogs } from '../../../src/toolbar/dialogs/shell.js';
import type { ApplyRibbonCommandDeps } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import {
  dispatchMacRibbonCommand,
  isMacRibbonCommandSupported,
} from '../../../src/toolbar/ribbon/mac/dispatch.js';
import {
  MAC_RIBBON_COMMAND_IDS,
  MAC_RIBBON_DISABLED_COMMAND_IDS,
} from '../../../src/toolbar/ribbon/mac/model.js';

const reported = vi.hoisted(() => [] as string[]);

vi.mock('../../../src/toolbar/dialogs/report.js', async (importOriginal) => {
  const original = await importOriginal<typeof import('../../../src/toolbar/dialogs/report.js')>();
  return {
    ...original,
    showReport: (opts: { items: readonly { label: string }[] }) => {
      for (const item of opts.items) reported.push(item.label);
      return Promise.resolve();
    },
  };
});

afterEach(() => {
  cancelOpenAppDialogs();
  document.body.replaceChildren();
  reported.length = 0;
});

/** Instance whose unknown members are inert functions, so only dispatch routing is exercised. */
const stubInstance = (): SpreadsheetInstance => {
  const host = document.createElement('div');
  document.body.append(host);
  const base: Record<string, unknown> = {
    store: createSpreadsheetStore(),
    host,
    i18n: createI18nController({ locale: 'en' }),
    commands: { policy: undefined },
  };
  return new Proxy(base, {
    get: (target, prop) => {
      if (prop in target) return target[prop as string];
      if (typeof prop === 'symbol' || prop === 'then') return undefined;
      return () => undefined;
    },
  }) as unknown as SpreadsheetInstance;
};

describe('Mac ribbon dispatch coverage', () => {
  it('surfaces a failing handler through the mocked report dialog', async () => {
    const deps = {
      inst: stubInstance(),
      runtime: { projectFormatToolbar: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    dispatchMacRibbonCommand('mac.draw.toggle', deps, () => true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(reported).toContain('Ink is unavailable for this spreadsheet.');
  });

  it('routes every supported, enabled command to a handler', async () => {
    const instance = stubInstance();
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    const ids = MAC_RIBBON_COMMAND_IDS.filter(
      (id) => isMacRibbonCommandSupported(id) && !MAC_RIBBON_DISABLED_COMMAND_IDS.has(id),
    );
    expect(ids.length).toBeGreaterThan(100);
    for (const id of ids) {
      reported.length = 0;
      dispatchMacRibbonCommand(id, deps, () => true);
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(
        reported.filter((label) => label.includes('Unavailable ribbon command')),
        id,
      ).toEqual([]);
    }
  });
});
