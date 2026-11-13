import { afterEach, describe, expect, it } from 'vitest';
import { keepSheetView, newSheetView } from '../../../src/commands/sheet-views.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';
import { macRibbonMenuIdForCommand } from '../../../src/toolbar/ribbon/mac/menu-manifest.js';
import { createMacRibbonMenuFactory } from '../../../src/toolbar/ribbon/mac/menus.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { stubHelpers } from '../mount/toolbar/fixtures.js';

const viewTab = () =>
  buildRibbonModel('ja', { profile: 'excel365Mac' }).find((tab) => tab.id === 'view');

/** A menu's items as id and drawn text, with whether each is greyed out. */
const menuItems = (commandId: string) =>
  Array.from(
    createMacRibbonMenuFactory('ja')(commandId)?.querySelectorAll<HTMLElement>(
      '[data-ribbon-command]',
    ) ?? [],
  ).map((item) => [
    item.dataset.ribbonCommand,
    item.textContent,
    item.getAttribute('aria-disabled') === 'true',
  ]);

describe('Mac View tab', () => {
  it('lays out the groups and their commands in the desktop order', () => {
    expect(viewTab()?.groups.map((group) => group.commands.map((command) => command.id))).toEqual([
      [
        'sheetViewSelect',
        'mac.view.sheetViewKeep',
        'mac.view.sheetViewExit',
        'mac.view.sheetViewNew',
        'mac.view.sheetViewOptions',
      ],
      [
        'mac.view.standard',
        'mac.view.pageBreakPreview',
        'mac.view.pageLayout',
        'mac.view.customViews',
      ],
      [
        'mac.view.navigation',
        'mac.view.ruler',
        'mac.view.gridlines',
        'mac.view.formulaBar',
        'mac.view.headings',
        'mac.view.dataTypeIcons',
        'mac.view.cellFocus',
      ],
      ['mac.view.zoom', 'mac.view.zoom100', 'mac.view.zoomSelection'],
      [
        'mac.view.newWindow',
        'mac.view.arrange',
        'mac.view.freeze',
        'mac.view.split',
        'mac.view.hide',
        'mac.view.unhide',
        'mac.view.sideBySide',
        'mac.view.syncScroll',
        'mac.view.resetWindow',
        'mac.view.switchWindow',
      ],
      ['mac.view.macros', 'mac.view.recordMacro', 'mac.view.relativeRecording'],
    ]);
  });

  it('opens Focus Cell as a menu of its three entries, not built yet', () => {
    expect(macRibbonMenuIdForCommand('mac.view.cellFocus')).toBe('menu-mac-view-cell-focus');
    expect(menuItems('mac.view.cellFocus')).toEqual([
      ['mac.view.cellFocus.focusCell', 'セルにフォーカス', true],
      ['mac.view.cellFocus.color', 'フォーカス セルの色', true],
      ['mac.view.cellFocus.autoHighlight', '自動強調表示を表示する', true],
    ]);
  });

  it('labels the Freeze Panes button and its menu entry as each is drawn', () => {
    expect(macRibbonLabelForCommand('mac.view.freeze', 'ja')).toBe('ウインドウ枠の固定');
    expect(macRibbonLabelForCommand('mac.view.freeze', 'de')).toBe('Bereiche einfrieren');
    expect(menuItems('mac.view.freeze')).toEqual([
      ['mac.view.freeze.panes', 'ウィンドウ枠の固定', false],
      ['mac.view.freeze.firstRow', '先頭行の固定', false],
      ['mac.view.freeze.firstColumn', '先頭列の固定', false],
    ]);
    expect(macRibbonLabelForCommand('mac.view.freeze.panes', 'de')).toBe('Fenster fixieren');
  });

  it('labels the Sheet View group from the desktop strings', () => {
    const commands = viewTab()?.groups[0]?.commands ?? [];
    expect(commands.map((command) => command.label)).toEqual([
      'シート ビューの切り替え',
      '保持',
      '終了',
      '新規',
      'オプション',
    ]);
    expect(commands[0]?.options?.map((option) => option.label)).toEqual(['既定', '一時ビュー']);
    expect(
      ['sheetViewKeep', 'sheetViewExit', 'sheetViewNew'].map((id) =>
        macRibbonLabelForCommand(`mac.view.${id}`, 'de'),
      ),
    ).toEqual(['Beibehalten', 'Beenden', 'Neu']);
  });

  it('names the zoom commands as the desktop does', () => {
    expect(macRibbonLabelForCommand('mac.view.zoom100', 'ja')).toBe('ズーム (100%)');
    expect(macRibbonLabelForCommand('mac.view.zoom100', 'en')).toBe('Zoom to 100%');
    expect(macRibbonLabelForCommand('mac.view.zoomSelection', 'ja')).toBe(
      '選択範囲に合わせて拡大/縮小',
    );
  });

  describe('toggles', () => {
    let sheet: MountedStubSheet | undefined;
    afterEach(() => {
      sheet?.dispose();
      sheet = undefined;
    });

    it('shows Formula Bar pressed while the bar is shown, Ruler on, Zoom to Selection off', async () => {
      sheet = await mountStubSheet({
        workbook: await WorkbookHandle.createDefault({ preferStub: true }),
        locale: 'ja',
      });
      const { instance } = sheet;
      const host = document.createElement('div');
      const button = (id: string) => {
        const el = document.createElement('button');
        el.dataset.ribbonCommand = id;
        host.appendChild(el);
        return el;
      };
      const formulaBar = button('mac.view.formulaBar');
      const ruler = button('mac.view.ruler');
      const zoomSelection = button('mac.view.zoomSelection');
      projectMacRibbonState(host, instance);
      expect(formulaBar.getAttribute('aria-pressed')).toBe('true');
      expect(ruler.getAttribute('aria-pressed')).toBe('true');
      expect(zoomSelection.getAttribute('aria-pressed')).toBe('false');
      mutators.setFormulaBarVisible(instance.store, false);
      projectMacRibbonState(host, instance);
      expect(formulaBar.getAttribute('aria-pressed')).toBe('false');
    });

    it('offers Keep only for a temporary view and Exit only away from Default', async () => {
      sheet = await mountStubSheet({
        workbook: await WorkbookHandle.createDefault({ preferStub: true }),
        locale: 'ja',
      });
      const { instance } = sheet;
      const host = document.createElement('div');
      const button = (id: string) => {
        const el = document.createElement('button');
        el.dataset.ribbonCommand = id;
        host.appendChild(el);
        return el;
      };
      const keep = button('mac.view.sheetViewKeep');
      const exit = button('mac.view.sheetViewExit');
      const offered = () => [keep.disabled, exit.disabled].map((off) => !off);

      projectMacRibbonState(host, instance);
      expect(offered()).toEqual([false, false]);
      newSheetView(instance.store);
      projectMacRibbonState(host, instance);
      expect(offered()).toEqual([true, true]);
      keepSheetView(instance.store, 'view-0-1', 'View 1');
      projectMacRibbonState(host, instance);
      expect(offered()).toEqual([false, true]);
    });

    it('routes New, Keep and Exit from the ribbon to the sheet views', async () => {
      sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
      const { instance } = sheet;
      const host = document.createElement('div');
      document.body.appendChild(host);
      const tb = Spreadsheet.mountToolbar(host, instance, { helpers: stubHelpers() });
      const views = () => instance.store.getState().sheetViews;

      expect(tb.applyCommand('mac.view.sheetViewNew')).toBe(true);
      expect(views().activeViewId).toBe('temporary');
      expect(tb.applyCommand('mac.view.sheetViewKeep')).toBe(true);
      expect(views().views.map((view) => [view.id, view.name])).toEqual([['view-0-1', 'ビュー 1']]);
      expect(tb.applyCommand('mac.view.sheetViewExit')).toBe(true);
      expect(views().activeViewId).toBeNull();
      tb.dispose();
      host.remove();
    });
  });
});
