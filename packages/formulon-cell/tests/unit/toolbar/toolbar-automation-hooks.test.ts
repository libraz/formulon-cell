import '../../api/setup-dom.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { mount, type SpreadsheetInstance } from '../../../src/index.js';
import { mountToolbar } from '../../../src/toolbar/index.js';

const makeHost = (): HTMLElement => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  return host;
};

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const i of live.splice(0)) i.dispose();
});

// The dropdowns are built on open; a menu of the same shape stands in for the open one.
const clickItem = (menuId: string, attr: string, value: string): void => {
  const menu = document.createElement('div');
  menu.id = menuId;
  menu.className = 'fc-tb__menu';
  const item = document.createElement('button');
  item.dataset[attr] = value;
  menu.append(item);
  document.body.append(menu);
  item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  menu.remove();
};

describe('ribbon automation hooks', () => {
  it('routes the custom script and add-in manager items to the host hooks', async () => {
    const instance = await mount(makeHost());
    live.push(instance);
    const runScript = vi.fn();
    const addInManager = vi.fn();
    const toolbar = await mountToolbar(makeHost(), instance, {
      hooks: { automation: { runScript, addInManager } },
    });
    clickItem('menu-script', 'scriptAction', 'custom');
    expect(runScript).toHaveBeenCalledOnce();
    clickItem('menu-add-ins', 'addInAction', 'manage');
    expect(addInManager).toHaveBeenCalledOnce();
    toolbar.dispose();
  });

  it('leaves the other script items on the default path', async () => {
    const instance = await mount(makeHost());
    live.push(instance);
    const runScript = vi.fn();
    const toolbar = await mountToolbar(makeHost(), instance, {
      hooks: { automation: { runScript } },
    });
    instance.applyChanges([{ at: 'A1', input: 'x' }], { history: 'record' });
    expect(JSON.stringify(instance.workbook.sheet(0).getValue('A1'))).toContain('x');
    clickItem('menu-script', 'scriptAction', 'clear');
    expect(runScript).not.toHaveBeenCalled();
    expect(JSON.stringify(instance.workbook.sheet(0).getValue('A1'))).not.toContain('x');
    toolbar.dispose();
  });
});
