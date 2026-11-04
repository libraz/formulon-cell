import { afterEach, describe, expect, it } from 'vitest';
import { quickAnalysis } from '../../../src/extensions/built-ins.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

// The Mac app has no Quick Analysis: no badge, no shortcut, no command.
describe('Quick Analysis by platform profile', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  const quick = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-quick');
  const pressCtrlQ = (host: HTMLElement): void => {
    const grid = host.querySelector<HTMLElement>('.fc-host__grid') ?? host;
    grid.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'q', ctrlKey: true, bubbles: true, cancelable: true }),
    );
  };

  it('is offered on the default profile', async () => {
    sheet = await mountStubSheet({ ui: { platform: 'default' } });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(quick()).not.toBeNull();
    sheet.instance.openQuickAnalysis();
    expect(quick()?.hidden).toBe(false);
  });

  it('is absent on the Mac profile: no popover, command or shortcut', async () => {
    sheet = await mountStubSheet({ ui: { platform: 'mac' } });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(quick()).toBeNull();
    expect(document.querySelector('.fc-quick__button')).toBeNull();
    sheet.instance.openQuickAnalysis();
    pressCtrlQ(sheet.host);
    expect(quick()).toBeNull();
  });

  it('goes away and comes back when the profile switches', async () => {
    sheet = await mountStubSheet({ ui: { platform: 'default' } });
    expect(quick()).not.toBeNull();
    sheet.instance.setUi({ platform: 'mac' });
    expect(quick()).toBeNull();
    sheet.instance.setUi({ platform: 'default' });
    expect(quick()).not.toBeNull();
  });

  it('stays absent on the Mac profile when composed from the public extension', async () => {
    sheet = await mountStubSheet({
      ui: { platform: 'mac' },
      features: { quickAnalysis: false },
      extensions: [quickAnalysis()],
    });
    expect(quick()).toBeNull();
    sheet.instance.openQuickAnalysis();
    expect(quick()).toBeNull();
  });
});
