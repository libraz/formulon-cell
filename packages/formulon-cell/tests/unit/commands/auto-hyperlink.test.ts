import { describe, expect, it, vi } from 'vitest';
import {
  detectTypedWebUrl,
  getTypedWebUrlFormatPatch,
  resolveHyperlinkStyleXf,
} from '../../../src/commands/auto-hyperlink.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;
const B1 = { sheet: 0, row: 0, col: 1 } as const;
const URL = 'https://example.com/path?q=1#fragment';

describe('detectTypedWebUrl', () => {
  it.each([
    'https://example.com/path?query=1#section',
    'http://example.com/',
    'HTTPS://EXAMPLE.COM/Path',
    'https://[2001:db8::1]/path',
  ])('accepts the whole absolute web URL without rewriting %s', (raw) => {
    expect(detectTypedWebUrl(raw)).toBe(raw);
  });

  it.each([
    "'https://example.com",
    ' https://example.com',
    'https://example.com/path here',
    '=HYPERLINK("https://example.com")',
    'mailto:person@example.com',
    'tel:+15555550123',
    'javascript:alert(1)',
    'https://example.com.invalid space',
    'prefix https://example.com',
    'example.com',
    'https://',
    'http://[not-an-ip]',
  ])('rejects non-typed web URL input %s', (raw) => {
    expect(detectTypedWebUrl(raw)).toBeNull();
  });
});

const addNamedStyleWithFont = (
  wb: WorkbookHandle,
  name: string,
  builtinId: number | null,
  options: { size: number; colorArgb: number; colorKind?: number; hidden?: boolean },
) => {
  const baseXf = wb.getCellStyleXf(0);
  const baseFont = baseXf ? wb.getFontRecord(baseXf.fontIndex) : null;
  expect(baseXf).not.toBeNull();
  expect(baseFont).not.toBeNull();
  if (!baseXf || !baseFont) throw new Error('default style XF/font is unavailable');

  const fontIndex = wb.addFontRecord({
    ...baseFont,
    size: options.size,
    colorArgb: options.colorArgb,
    bold: false,
    italic: false,
    strike: false,
    underline: 0,
    vertAlign: 0,
    color: {
      kind: options.colorKind ?? 1,
      rgb: options.colorArgb,
      theme: 0,
      tint: 0,
      indexed: 0,
    },
  });
  const xfId = wb.addCellStyleXfRecord({ ...baseXf, fontIndex });
  expect(fontIndex).toBeGreaterThanOrEqual(0);
  expect(xfId).toBeGreaterThanOrEqual(0);
  expect(
    wb.setNamedCellStyleRecord({
      name,
      xfId,
      builtinId,
      hidden: options.hidden ?? false,
      customBuiltin: true,
    }),
  ).toBe(true);
  return { xfId, xf: wb.getCellStyleXf(xfId) };
};

describe('resolveHyperlinkStyleXf', () => {
  it.each([true, false])(
    'prefers a localized built-in Hyperlink over an earlier custom name (hidden=%s)',
    async (hidden) => {
      const wb = await WorkbookHandle.createDefault();
      try {
        expect(wb.isStub).toBe(false);
        const custom = addNamedStyleWithFont(wb, 'Hyperlink', null, {
          size: 11,
          colorArgb: 0xffcc0000,
        });
        const builtin = addNamedStyleWithFont(wb, 'LienWeb', 8, {
          size: 19,
          colorArgb: 0xff0055cc,
          hidden,
        });

        expect(custom.xfId).not.toBe(builtin.xfId);
        const resolved = resolveHyperlinkStyleXf(wb);
        expect(resolved).toMatchObject({
          name: 'LienWeb',
          xfId: builtin.xfId,
          builtinId: 8,
          xf: builtin.xf,
        });
        expect(wb.getFontRecord(resolved?.xf.fontIndex ?? -1)).toMatchObject({
          size: 19,
          colorArgb: 0xff0055cc,
        });
        expect(
          Array.from({ length: wb.cellStyleCount() }, (_, index) => wb.getCellStyle(index)).find(
            (style) => style?.name === 'LienWeb',
          ),
        ).toMatchObject({ hidden, builtinId: 8 });
      } finally {
        wb.dispose();
      }
    },
  );

  it('returns the complete XF from a customized built-in Hyperlink style', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const customized = addNamedStyleWithFont(wb, 'Link Style', 8, {
        size: 16,
        colorArgb: 0xff008844,
      });
      const resolved = resolveHyperlinkStyleXf(wb);
      expect(resolved).toMatchObject({
        name: 'Link Style',
        xfId: customized.xfId,
        builtinId: 8,
        xf: customized.xf,
      });
      expect(wb.getFontRecord(resolved?.xf.fontIndex ?? -1)).toMatchObject({
        size: 16,
        colorArgb: 0xff008844,
      });
    } finally {
      wb.dispose();
    }
  });

  it('uses a custom Hyperlink name when there is no built-in id 8', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const custom = addNamedStyleWithFont(wb, '  Hyperlink  ', null, {
        size: 13,
        colorArgb: 0xffaa00aa,
      });
      const resolved = resolveHyperlinkStyleXf(wb);
      expect(resolved).toMatchObject({
        name: '  Hyperlink  ',
        xfId: custom.xfId,
        xf: custom.xf,
      });
      expect([0xffffffff, -1]).toContain(resolved?.builtinId);
    } finally {
      wb.dispose();
    }
  });

  it.each([
    { name: 'Hyperlink', builtinId: 9 },
    { name: 'FollowedHyperlink', builtinId: 9 },
  ])('does not treat $name with built-in id $builtinId as the Hyperlink style', async (style) => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      addNamedStyleWithFont(wb, style.name, style.builtinId, {
        size: 14,
        colorArgb: 0xff336699,
      });
      expect(resolveHyperlinkStyleXf(wb)).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('returns null when style access is unsupported or an XF cannot be resolved', () => {
    const unsupported = {
      capabilities: { cellStyles: false },
      cellStyleCount: vi.fn(() => 1),
      getCellStyle: vi.fn(() => null),
      getCellStyleXf: vi.fn(() => null),
    };
    expect(resolveHyperlinkStyleXf(unsupported as never)).toBeNull();
    expect(unsupported.cellStyleCount).not.toHaveBeenCalled();

    const missingStyle = {
      capabilities: { cellStyles: true },
      cellStyleCount: vi.fn(() => 1),
      getCellStyle: vi.fn(() => null),
      getCellStyleXf: vi.fn(() => null),
    };
    expect(resolveHyperlinkStyleXf(missingStyle as never)).toBeNull();
    expect(missingStyle.cellStyleCount).toHaveBeenCalledOnce();

    const invalidXfId = {
      capabilities: { cellStyles: true },
      cellStyleCount: vi.fn(() => 1),
      getCellStyle: vi.fn(() => ({
        name: 'Hyperlink',
        xfId: -1,
        builtinId: 0xffffffff,
      })),
      getCellStyleXf: vi.fn(() => null),
    };
    expect(resolveHyperlinkStyleXf(invalidXfId as never)).toBeNull();
    expect(invalidXfId.getCellStyleXf).not.toHaveBeenCalled();
  });

  it.each([true, false])(
    'materializes all font fields and preserves other formats for a customized built-in style (hidden=%s)',
    async (hidden) => {
      const wb = await WorkbookHandle.createDefault();
      const store = createSpreadsheetStore();
      try {
        expect(wb.isStub).toBe(false);
        const linkStyle = addNamedStyleWithFont(wb, 'LienWeb', 8, {
          size: 24,
          colorArgb: 0xff000000,
          hidden,
        });
        expect(linkStyle.xf).not.toBeNull();
        wb.setText(A1, URL);
        const existing = {
          fontFamily: 'Calibri',
          fontSize: 32,
          bold: true,
          italic: true,
          strike: true,
          underline: 'double' as const,
          fontVertAlign: 'superscript' as const,
          fontOutline: true,
          fontShadow: true,
          fontCondense: true,
          fontExtend: true,
          color: '#123456',
          numFmt: { kind: 'custom' as const, pattern: '0.000' },
          fill: '#eeeeee',
          borders: { top: { style: 'thin' as const, color: '#abcdef' } },
          locked: false,
          comment: 'keep note',
          validation: { kind: 'list' as const, source: ['A', 'B'] },
          hyperlink: 'https://old.example/',
          hyperlinkLocation: 'OldSheet!A1',
          hyperlinkDisplay: 'old label',
          hyperlinkTooltip: 'old tip',
          hyperlinkSpan: { lastRow: A1.row, lastCol: A1.col },
        };
        mutators.setCellFormat(store, A1, existing);

        const patch = getTypedWebUrlFormatPatch(wb, store.getState(), A1, URL);
        expect(patch).toMatchObject({
          fontFamily: wb.getResolvedFontRecord(linkStyle.xf?.fontIndex ?? -1)?.name,
          fontSize: 32,
          bold: false,
          italic: false,
          strike: false,
          underline: false,
          color: '#000000',
          hyperlink: URL,
        });
        expect(patch).toHaveProperty('fontVertAlign', undefined);
        expect(patch).toMatchObject({
          fontOutline: false,
          fontShadow: false,
          fontCondense: false,
          fontExtend: false,
        });

        mutators.setCellFormat(store, A1, patch);
        const after = store.getState().format.formats.get(addrKey(A1));
        expect(after).toMatchObject({
          numFmt: existing.numFmt,
          fill: existing.fill,
          borders: existing.borders,
          locked: false,
          comment: existing.comment,
          validation: existing.validation,
          hyperlink: URL,
        });
        expect(after?.hyperlinkLocation).toBeUndefined();
        expect(after?.hyperlinkDisplay).toBeUndefined();
        expect(after?.hyperlinkTooltip).toBeUndefined();
        expect(after?.hyperlinkSpan).toBeUndefined();
      } finally {
        wb.dispose();
      }
    },
  );

  it('uses workbook font 0 and theme hyperlink colour when no named style exists', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.isStub).toBe(false);
      const existingHyperlinks = Array.from({ length: wb.cellStyleCount() }, (_, index) =>
        wb.getCellStyle(index),
      ).filter((style) => style?.builtinId === 8);
      for (const style of existingHyperlinks) {
        if (style) expect(wb.removeNamedCellStyle(style.name)).toBe(true);
      }
      wb.setText(A1, URL);
      mutators.setCellFormat(store, A1, { fontFamily: 'Arial', fontSize: 40, underline: false });

      const patch = getTypedWebUrlFormatPatch(wb, store.getState(), A1, URL);
      const displayFont = wb.workbookDisplayFont;
      const hyperlinkColor = wb.getTheme()?.colors[10];
      expect(displayFont).not.toBeNull();
      expect(hyperlinkColor).toBeTypeOf('number');
      expect(patch).toMatchObject({
        fontFamily: displayFont?.name,
        fontSize: displayFont ? (displayFont.size * 4) / 3 : undefined,
        underline: 'single',
        color:
          hyperlinkColor === undefined
            ? undefined
            : `#${(hyperlinkColor & 0xffffff).toString(16).padStart(6, '0')}`,
        hyperlink: URL,
      });
    } finally {
      wb.dispose();
    }
  });

  it('materializes resolved none/automatic font colour as explicit black', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      const style = addNamedStyleWithFont(wb, 'LienWeb', 8, {
        size: 11,
        colorArgb: 0,
        colorKind: 0,
      });
      wb.setText(A1, URL);
      const resolved = wb.getResolvedFontRecord(style.xf?.fontIndex ?? -1);
      expect(resolved?.color?.kind).toBe(0);
      expect(resolved?.colorArgb).toBe(0);
      expect(getTypedWebUrlFormatPatch(wb, store.getState(), A1, URL)?.color).toBe('#000000');
    } finally {
      wb.dispose();
    }
  });

  it('skips an incomplete named-style font, existing multi-cell links, formulas and locked cells', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(wb.isStub).toBe(false);
      addNamedStyleWithFont(wb, 'LienWeb', 8, { size: 13, colorArgb: 0xff0055cc });
      wb.setText(A1, URL);
      const originalGetFont = wb.getResolvedFontRecord.bind(wb);
      vi.spyOn(wb, 'getResolvedFontRecord').mockImplementation((index) =>
        index === resolveHyperlinkStyleXf(wb)?.xf.fontIndex ? null : originalGetFont(index),
      );
      expect(getTypedWebUrlFormatPatch(wb, store.getState(), A1, URL)).toBeNull();
      vi.restoreAllMocks();

      mutators.setCellFormat(store, A1, {
        hyperlink: 'https://anchor.example/',
        hyperlinkSpan: { lastRow: 1, lastCol: 1 },
      });
      wb.setText(B1, URL);
      expect(getTypedWebUrlFormatPatch(wb, store.getState(), B1, URL)).toBeNull();
      expect(getTypedWebUrlFormatPatch(wb, store.getState(), A1, URL)).toBeNull();

      mutators.setCellFormat(store, A1, null);
      mutators.setCellFormat(store, B1, null);
      wb.setFormula(B1, '=1');
      expect(getTypedWebUrlFormatPatch(wb, store.getState(), B1, URL)).toBeNull();

      wb.setText(B1, URL);
      mutators.setSheetProtected(store, 0, true, {
        permissions: { formatCells: true, insertHyperlinks: true },
      });
      expect(getTypedWebUrlFormatPatch(wb, store.getState(), B1, URL)).toBeNull();
    } finally {
      wb.dispose();
    }
  });
});
