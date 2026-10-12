import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { formatPainter } from '../../src/extensions/built-ins.js';
import type { FormatPainterHandle } from '../../src/interact/format-painter.js';
import { cellRectUnclamped, hitTest, layoutForView } from '../../src/render/geometry.js';
import { mountStubSheet } from '../test-utils/mount.js';

describe('mounted Format Painter inherits native column fonts', () => {
  for (const extension of [false, true]) {
    it.each(['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'])(
      `%s (extension=${extension})`,
      async (locale) => {
        const workbook = await WorkbookHandle.createDefault({ locale });
        const xf = workbook.getCellXf(0);
        const font = xf && workbook.getFontRecord(xf.fontIndex);
        if (!xf || !font) throw new Error('missing real workbook Normal font');
        const fontIndex = workbook.addFontRecord({ ...font, name: 'Consolas', size: 22 });
        const xfIndex = workbook.addXfRecord({ ...xf, fontIndex, applyFont: true });
        workbook.setText({ sheet: 0, row: 0, col: 2 }, 'destination');
        const sheet = await mountStubSheet({
          workbook,
          locale,
          ...(extension
            ? { features: { formatPainter: false }, extensions: [formatPainter()] }
            : {}),
        });
        try {
          sheet.instance.store.setState((state) => ({
            ...state,
            viewport: {
              ...state.viewport,
              widthPx: 800,
              heightPx: 600,
              rowCount: 25,
              colCount: 10,
            },
          }));
          expect(workbook.setColumnStyle(0, 0, 0, xfIndex)).toBe(true);
          const inherited = workbook.getEffectiveStyle(0, 0, 0);
          const inheritedXf = inherited && workbook.getCellXf(inherited.xfIndex);
          expect(inheritedXf && workbook.getFontRecord(inheritedXf.fontIndex)?.name).toBe(
            'Consolas',
          );
          expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
          expect(workbook.isStub).toBe(false);
          const handle = (
            extension ? sheet.instance.features.formatPainter : sheet.instance.formatPainter
          ) as FormatPainterHandle | undefined;
          if (!handle) throw new Error('Format Painter was not mounted');
          handle.activate();
          expect(handle.isActive()).toBe(true);
          const state = sheet.instance.store.getState();
          const rect = cellRectUnclamped(layoutForView(state), state.viewport, 0, 2);
          expect(
            hitTest(layoutForView(state), state.viewport, rect.x + rect.w / 2, rect.y + rect.h / 2),
          ).toEqual({ row: 0, col: 2 });
          const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
          const box = grid?.getBoundingClientRect();
          if (!box) throw new Error('grid was not mounted');
          for (const type of ['pointerdown', 'pointerup']) {
            sheet.host.dispatchEvent(
              new PointerEvent(type, {
                clientX: box.x + rect.x + rect.w / 2,
                clientY: box.y + rect.y + rect.h / 2,
                button: 0,
                pointerId: 1,
                bubbles: true,
                cancelable: true,
              }),
            );
          }
          const formats = sheet.instance.store.getState().format.formats;
          expect(
            formats.get('0:0:0'),
            JSON.stringify(sheet.instance.store.getState().selection),
          ).toBeUndefined();
          expect(
            formats.get('0:0:2'),
            JSON.stringify(sheet.instance.store.getState().selection),
          ).toMatchObject({
            fontFamily: 'Consolas',
            fontSize: 22 * (96 / 72),
          });
          expect(workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
            kind: 'text',
            value: 'destination',
          });
          expect(sheet.instance.undo()).toBe(true);
          expect(sheet.instance.store.getState().format.formats.get('0:0:2')).toBeUndefined();
        } finally {
          sheet.dispose();
          workbook.dispose();
        }
      },
    );
  }
});
