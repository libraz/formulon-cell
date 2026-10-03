import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { fireContextMenu, newWb, root, visibleMenu } from './fixtures.js';

describe('attachContextMenu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;
  let onAfterCommit: Mock<() => void>;
  let unregisterController: (() => void) | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
    detach = null;
    unregisterController = null;
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    unregisterController?.();
    document.body.innerHTML = '';
    vi.restoreAllMocks();
  });

  describe('teardown', () => {
    it('detach removes the menu from the DOM', () => {
      detach = attachContextMenu({ host, store, wb });
      expect(document.querySelector('.fc-ctxmenu')).not.toBeNull();
      detach();
      detach = null;
      expect(document.querySelector('.fc-ctxmenu')).toBeNull();
    });

    it('after detach, contextmenu events do not open the menu', () => {
      detach = attachContextMenu({ host, store, wb });
      detach();
      detach = null;
      fireContextMenu(host, 200, 70);
      expect(visibleMenu()).toBeNull();
    });
  });

  describe('DOM primitives', () => {
    it('keeps context menu button DOM on the shared interaction primitive', () => {
      const source = ['src/interact/context-menu.ts', 'src/interact/context-menu-entries.ts']
        .map((path) => readFileSync(join(root, path), 'utf8'))
        .join('\n');

      expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
      expect(source).toContain('const createContextMenuItemButton');
      expect(source).toContain('const createContextSubmenuButton');
      expect(source).toContain('const createContextMiniToolbarButton');
      expect(source).toContain('const btn = createContextSubmenuButton(entry)');
      expect(source).toContain('const btn = createContextMenuItemButton(entry)');
      expect(source).toContain('const btn = createContextMiniToolbarButton(item)');
      expect(source).not.toContain("document.createElement('button')");
      expect(source).not.toContain("arrow.textContent = '›'");
    });

    it('renders context submenu chevrons as fixed CSS glyphs, not text arrows', () => {
      const css = readFileSync(join(root, 'src/styles/core/app/overlays/context-menu.css'), 'utf8');
      detach = attachContextMenu({ host, store, wb, onAfterCommit });

      fireContextMenu(host, 200, 70);

      const arrow = document.querySelector<HTMLElement>('.fc-ctxmenu__arrow');
      expect(arrow).toBeTruthy();
      expect(arrow?.textContent).toBe('');
      expect(arrow?.getAttribute('aria-hidden')).toBe('true');
      expect(css).toMatch(/\.fc-ctxmenu__arrow\s*\{[\s\S]*?width: 8px;[\s\S]*?height: 12px;/);
      expect(css).toMatch(
        /\.fc-ctxmenu__arrow::before\s*\{[\s\S]*?border: solid currentColor;[\s\S]*?transform: translateY\(-50%\) rotate\(45deg\);/,
      );
    });

    it('uses colored Excel-like SVG backgrounds for high-frequency context menu icons', () => {
      const css = readFileSync(join(root, 'src/styles/core/app/overlays/context-menu.css'), 'utf8');

      expect(css).toMatch(
        /\.fc-ctxmenu\s*\{[\s\S]*?padding: 2px 0;[\s\S]*?border-radius: 2px;[\s\S]*?font-size: 12px;/,
      );
      expect(css).toMatch(
        /\.fc-ctxmenu\s*\{[\s\S]*?0 8px 24px rgba\(0, 0, 0, 0\.18\),[\s\S]*?0 1px 3px rgba\(0, 0, 0, 0\.14\);/,
      );
      expect(css).toContain('--fc-menu-icon-image');
      expect(css).toMatch(
        /\.fc-ctxmenu__item\s*\{[\s\S]*?grid-template-columns: 26px minmax\(0, 1fr\) auto;[\s\S]*?min-height: 26px;[\s\S]*?padding: 3px 18px 3px 12px;/,
      );
      expect(css).toMatch(
        /\.fc-ctxmenu__item::before\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;[\s\S]*?background-size: 18px 18px;/,
      );
      expect(css).toMatch(
        /\.fc-ctxmenu__item\[data-fc-action="copy"\]::before,[\s\S]*?\.fc-ctxmenu__item\[data-fc-action="sort"\]::before,[\s\S]*?\{[\s\S]*?-webkit-mask: none;[\s\S]*?mask: none;/,
      );
      const cutIconCss =
        /\.fc-ctxmenu__item\[data-fc-action="cut"\]\s*\{(?<body>[\s\S]*?)\n {2}\}/.exec(css)?.groups
          ?.body ?? '';
      expect(cutIconCss).toContain("cx='5.2' cy='15.2' r='2.35' fill='%232f75b5'");
      expect(cutIconCss).toContain("cx='14.8' cy='15.2' r='2.35' fill='%232f75b5'");
      expect(cutIconCss).toContain("cx='5.2' cy='15.2' r='1.12' fill='%23fff'");
      expect(cutIconCss).toContain("cx='14.8' cy='15.2' r='1.12' fill='%23fff'");
      expect(cutIconCss).toContain("cx='10' cy='10.7'");
      expect(cutIconCss).toContain('%232f75b5');
      expect(cutIconCss).toContain('%23107c41');
      expect(css).toMatch(
        /\.fc-ctxmenu__item\[data-fc-action="formatCells"\]\s*\{[\s\S]*?%23107c41[\s\S]*?%232f75b5/,
      );
      expect(css).toMatch(/\.fc-ctxmenu__item\[data-fc-action="clear"\],[\s\S]*?%23c00000/);
    });

    it('uses colored SVG backgrounds for remaining semantic context menu row icons', () => {
      const css = readFileSync(join(root, 'src/styles/core/app/overlays/context-menu.css'), 'utf8');
      const actions = [
        'pasteAll',
        'pasteFormulas',
        'pasteFormulasNumFmt',
        'pasteValues',
        'pasteValuesNumFmt',
        'pasteFormatsOnly',
        'pasteTranspose',
        'insertCopiedCells',
        'rowInsertAbove',
        'rowInsertBelow',
        'colInsertLeft',
        'colInsertRight',
        'rowDelete',
        'colDelete',
        'rowHide',
        'colHide',
        'rowUnhide',
        'colUnhide',
        'rowHeight',
        'colWidth',
        'defineName',
        'filterClear',
        'filterReapply',
        'filterByValue',
        'sortAsc',
        'sortDesc',
        'selectAll',
        'insertComment',
        'insertHyperlink',
        'toggleWatch',
      ];

      for (const action of actions) {
        expect(css).toContain(`.fc-ctxmenu__item[data-fc-action="${action}"]::before`);
        expect(css).toMatch(
          new RegExp(
            `\\.fc-ctxmenu__item\\[data-fc-action="${action}"\\][\\s\\S]*?--fc-menu-icon-image: url`,
          ),
        );
      }
      expect(css).toMatch(/\.fc-ctxmenu__item\[data-fc-action="rowDelete"\],[\s\S]*?%23c00000/);
      expect(css).toMatch(/\.fc-ctxmenu__item\[data-fc-action="rowHide"\],[\s\S]*?%23c00000/);
      expect(css).toMatch(
        /\.fc-ctxmenu__item\[data-fc-action="insertHyperlink"\]\s*\{[\s\S]*?%232f75b5[\s\S]*?%23107c41/,
      );
    });

    it('uses colored SVG backgrounds for the context mini toolbar icons', () => {
      const css = readFileSync(join(root, 'src/styles/core/app/overlays/context-menu.css'), 'utf8');

      expect(css).toContain('--fc-mini-icon-image');
      expect(css).toMatch(
        /\.fc-ctxmenu__mini-btn\[data-fc-action="bold"\]::before,[\s\S]*?\.fc-ctxmenu__mini-btn\[data-fc-action="formatCells"\]::before\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;[\s\S]*?background-size: 18px 18px;/,
      );
      for (const action of [
        'bold',
        'italic',
        'underline',
        'alignLeft',
        'alignCenter',
        'alignRight',
        'borders',
        'formatCells',
      ]) {
        expect(css).toMatch(
          new RegExp(
            `\\.fc-ctxmenu__mini-btn\\[data-fc-action="${action}"\\]::before\\s*\\{[\\s\\S]*?--fc-mini-icon-image: url`,
          ),
        );
      }
      expect(css).toMatch(
        /\.fc-ctxmenu__mini-btn\[data-fc-action="underline"\]::before\s*\{[\s\S]*?%23107c41/,
      );
      expect(css).toMatch(
        /\.fc-ctxmenu__mini-btn\[data-fc-action="borders"\]::before\s*\{[\s\S]*?%23107c41/,
      );
      expect(css).toMatch(
        /\.fc-ctxmenu__mini-btn\[data-fc-action="formatCells"\]::before\s*\{[\s\S]*?%23107c41[\s\S]*?%232f75b5/,
      );
    });
  });
});
