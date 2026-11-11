import '../../api/setup-dom.js';
import { afterEach, describe, expect, it } from 'vitest';
import { mount, type SpreadsheetInstance } from '../../../src/index.js';
import { mountToolbar, ribbon } from '../../../src/toolbar/index.js';
import { expectFailure } from '../../api/helpers.js';

const makeHost = (): HTMLElement => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  return host;
};

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const i of live.splice(0)) i.dispose();
});

describe('toolbar object', () => {
  it('ribbon() carries its options under the names the ribbon mount reads', () => {
    expect(ribbon()).toEqual({ kind: 'ribbon' });
    expect(ribbon({ displayMode: 'tabsOnly', activeTab: 'data' })).toEqual({
      kind: 'ribbon',
      activeTab: 'data',
      ribbonDisplayMode: 'tabsOnly',
    });
  });

  it('opens an inline ribbon with the options ribbon() carries', async () => {
    const instance = await mount(makeHost(), { toolbar: ribbon({ activeTab: 'data' }) });
    live.push(instance);
    expect(instance.toolbar?.activeTab).toBe('data');
    instance.update({ toolbar: ribbon({ activeTab: 'insert' }) });
    expect(instance.toolbar?.activeTab).toBe('insert');
  });

  it('reports the inline ribbon state changes through toolbar.on', async () => {
    const instance = await mount(makeHost(), { toolbar: ribbon({ activeTab: 'home' }) });
    live.push(instance);
    const toolbar = instance.toolbar;
    const states: unknown[] = [];
    toolbar?.on('change', (state) => states.push(state.activeTab));
    toolbar?.update({ activeTab: 'data' });
    expect(states).toEqual(['data']);
    expect(instance.toolbar).toBe(toolbar);
    instance.update({ toolbar: null });
    expect(instance.toolbar).toBeNull();
  });

  it('mounts apart from the grid, reports state changes and leaves the instance untouched', async () => {
    const instance = await mount(makeHost());
    live.push(instance);
    const updates: unknown[] = [];
    instance.on('update', (e) => updates.push(e));
    const host = makeHost();
    const toolbar = await mountToolbar(host, instance, { activeTab: 'insert' });
    expect(instance.toolbar).toBeNull();
    expect(updates).toEqual([]);
    expect(toolbar.activeTab).toBe('insert');
    const seen: string[] = [];
    toolbar.on('change', (s) => seen.push(s.activeTab));
    toolbar.update({ activeTab: 'data' });
    expect(toolbar.activeTab).toBe('data');
    expect(seen).toContain('data');
    toolbar.dispose();
    expect(host.querySelector('[data-ribbon-tab]')).toBeNull();
  });

  it('rejects with the signal reason before touching the host', async () => {
    const instance = await mount(makeHost());
    live.push(instance);
    const host = makeHost();
    const reason = new Error('stop');
    const controller = new AbortController();
    controller.abort(reason);
    await expect(mountToolbar(host, instance, { signal: controller.signal })).rejects.toBe(reason);
    expect(host.childElementCount).toBe(0);
  });

  it('follows the instance theme and restores the host on dispose', async () => {
    const instance = await mount(makeHost(), { theme: 'ink' });
    live.push(instance);
    const host = makeHost();
    const toolbar = await mountToolbar(host, instance);
    expect(host.dataset.fcTheme).toBe('ink');
    instance.update({ theme: 'contrast' });
    expect(host.dataset.fcTheme).toBe('contrast');
    toolbar.dispose();
    expect(host.dataset.fcTheme).toBeUndefined();
  });

  it('execute throws policy-denied when the policy refuses the command', async () => {
    const instance = await mount(makeHost(), {
      policy: { operations: { insertRows: false } },
    });
    live.push(instance);
    const toolbar = await mountToolbar(makeHost(), instance);
    expectFailure(() => toolbar.execute('insertRows'), 'policy-denied');
    toolbar.dispose();
  });
});
