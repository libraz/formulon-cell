import '../../api/setup-dom.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { mount, type SpreadsheetInstance } from '../../../src/index.js';
import { mountToolbar, ribbon } from '../../../src/toolbar/index.js';

const makeHost = (): HTMLElement => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  return host;
};

const live: Array<{ dispose(): void }> = [];
afterEach(() => {
  for (const item of live.splice(0).reverse()) item.dispose();
  document.body.replaceChildren();
});

const clickScript = (host: HTMLElement): void => {
  host.querySelector<HTMLButtonElement>('[data-ribbon-tab="automate"]')?.click();
  const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="script"]');
  if (!button) throw new Error('Missing Script command');
  button.click();
};

const scriptMenuOpen = (): boolean => {
  const menu = document.querySelector<HTMLElement>('#menu-script');
  return !!menu && !menu.hidden;
};

describe('host-menu-first ribbon commands', () => {
  it('open their dropdown from a standalone toolbar once the host supplies the hook', async () => {
    const instance: SpreadsheetInstance = await mount(makeHost());
    live.push(instance);
    const host = makeHost();
    const runScript = vi.fn();
    live.push(
      await mountToolbar(host, instance, {
        ribbonTabs: ['home', 'automate'],
        hooks: { automation: { runScript } },
      }),
    );
    clickScript(host);
    expect(scriptMenuOpen()).toBe(true);
    expect(runScript).not.toHaveBeenCalled();
  });

  it('open their dropdown from the inline ribbon once the host supplies the hook', async () => {
    const host = makeHost();
    const runScript = vi.fn();
    const instance = await mount(host, {
      toolbar: ribbon({ ribbonTabs: ['home', 'automate'], hooks: { automation: { runScript } } }),
    });
    live.push(instance);
    clickScript(host);
    expect(scriptMenuOpen()).toBe(true);
    expect(runScript).not.toHaveBeenCalled();
  });

  it('run the built-in face action when the host supplies no hook', async () => {
    const host = makeHost();
    const instance = await mount(host, { toolbar: ribbon({ ribbonTabs: ['home', 'automate'] }) });
    live.push(instance);
    clickScript(host);
    expect(scriptMenuOpen()).toBe(false);
    expect(document.body.querySelector('.fc-tb__dlg')?.textContent).toContain('Script');
  });
});
