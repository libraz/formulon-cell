import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { syncTaskpaneDock } from '../../../src/interact/taskpane-dock.js';
import { attachWorkbookObjectsPanel } from '../../../src/interact/workbook-objects.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';
import { setupPalette } from './mac-formula-palette/fixtures.js';

describe('task-pane dock resize', () => {
  const fire = (target: EventTarget, type: string, x: number) =>
    target.dispatchEvent(
      new MouseEvent(type, { clientX: x, clientY: 100, button: 0, bubbles: true }),
    );
  const build = (platform = 'mac') => {
    const root = document.createElement('div');
    root.dataset.fcPlatform = platform;
    const dock = document.createElement('div');
    const pane = document.createElement('div');
    dock.appendChild(pane);
    root.appendChild(dock);
    document.body.appendChild(root);
    dock.getBoundingClientRect = () =>
      ({ left: 3539, right: 3840, width: 301, top: 0, bottom: 800, height: 800 }) as DOMRect;
    syncTaskpaneDock(dock);
    return dock;
  };
  afterEach(() => document.body.replaceChildren());

  it('follows the pointer from the left edge with no snapping', () => {
    const dock = build();
    fire(dock, 'pointerdown', 3540);
    fire(document, 'pointermove', 3471);
    expect(dock.style.flexBasis).toBe('370px');
    fire(document, 'pointermove', 3100);
    expect(dock.style.flexBasis).toBe('741px');
    fire(document, 'pointerup', 3100);
    fire(document, 'pointermove', 3000);
    expect(dock.style.flexBasis).toBe('741px');
  });

  it('stops at 271pt and 1001pt', () => {
    const dock = build();
    fire(dock, 'pointerdown', 3540);
    fire(document, 'pointermove', 3800);
    expect(dock.style.flexBasis).toBe('271px');
    fire(document, 'pointermove', 1000);
    expect(dock.style.flexBasis).toBe('1001px');
    fire(document, 'pointerup', 1000);
  });

  it('leaves presses away from the edge alone', () => {
    const dock = build();
    fire(dock, 'pointerdown', 3700);
    fire(document, 'pointermove', 3000);
    expect(dock.style.flexBasis).toBe('');
  });

  it('is a Mac-profile behaviour', () => {
    const dock = build('default');
    fire(dock, 'pointerdown', 3540);
    fire(document, 'pointermove', 3000);
    expect(dock.style.flexBasis).toBe('');
  });
});

describe('task-pane dock', () => {
  it('shows while any child pane is visible', () => {
    const dock = document.createElement('div');
    const a = document.createElement('div');
    const b = document.createElement('div');
    dock.append(a, b);
    a.hidden = true;
    b.hidden = true;
    syncTaskpaneDock(dock);
    expect(dock.hidden).toBe(true);
    b.hidden = false;
    syncTaskpaneDock(dock);
    expect(dock.hidden).toBe(false);
  });

  describe('shared by the formula palette and the PivotTable Fields pane', () => {
    let sheet: MountedStubSheet;
    beforeEach(async () => {
      sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault() });
    });
    afterEach(() => sheet.dispose());

    it('keeps the field pane docked when the palette closes', () => {
      const { dock, palette } = setupPalette(sheet);
      const wb = sheet.workbook;
      const rows: (string | number)[][] = [
        ['Region', 'Sales'],
        ['East', 12],
        ['West', 5],
      ];
      for (const [r, row] of rows.entries()) {
        for (const [c, value] of row.entries()) {
          const addr = { sheet: 0, row: r, col: c };
          if (typeof value === 'number') wb.setNumber(addr, value);
          else wb.setText(addr, value);
        }
      }
      const created = createPivotTableFromRange(wb, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        destination: { sheet: 0, row: 6, col: 0 },
        rowField: 'Region',
        valueField: 'Sales',
      });
      expect(created.ok).toBe(true);
      const objects = attachWorkbookObjectsPanel({ host: sheet.host, wb, strings: en, dock });

      palette.open();
      expect(objects.openPivotFieldList(0, 0)).toBe(true);
      const pane = dock.querySelector<HTMLElement>('.fc-objects--taskpane');
      expect(pane?.hidden).toBe(false);
      expect(dock.hidden).toBe(false);

      palette.close();
      expect(dock.hidden).toBe(false);
      expect(pane?.hidden).toBe(false);
      expect(objects.isPivotFieldListOpen()).toBe(true);

      objects.close();
      expect(dock.hidden).toBe(true);
      objects.detach();
      palette.detach();
    });
  });
});
