import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';

export const macStyleOracle = JSON.parse(
  readFileSync(
    resolve(
      dirname(fileURLToPath(import.meta.url)),
      '../../../fixtures/excel-mac-style-groups.json',
    ),
    'utf8',
  ),
) as {
  styles: readonly {
    name: string;
    galleryId: string | null;
    includedGroups: readonly string[];
  }[];
};

export const namedStyleWorkbook = (): WorkbookHandle =>
  ({
    getNamedCellStyles: () => [
      {
        index: 0,
        name: 'Normal',
        xfId: 0,
        builtinId: 0,
        iLevel: 0,
        customBuiltin: false,
      },
      {
        index: 1,
        name: 'Imported Review',
        xfId: 2,
        builtinId: -1,
        iLevel: 0,
        customBuiltin: false,
      },
    ],
    getCellStyleXf: (xfId: number) =>
      xfId === 2
        ? {
            fontIndex: 1,
            fillIndex: 1,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 2,
            verticalAlign: 1,
            wrapText: true,
          }
        : {
            fontIndex: 0,
            fillIndex: 0,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 0,
            verticalAlign: 2,
            wrapText: false,
          },
    getFontRecord: (fontIndex: number) => ({
      name: 'Calibri',
      size: 11,
      bold: fontIndex === 1,
      italic: false,
      strike: false,
      underline: 0,
      colorArgb: fontIndex === 1 ? 0xff006100 : 0xff000000,
    }),
    getFillRecord: (fillIndex: number) => ({
      pattern: fillIndex === 1 ? 1 : 0,
      fgArgb: fillIndex === 1 ? 0xffc6efce : 0,
      bgArgb: 0,
    }),
    getBorderRecord: () => ({
      left: { style: 0, colorArgb: 0xff000000 },
      right: { style: 0, colorArgb: 0xff000000 },
      top: { style: 0, colorArgb: 0xff000000 },
      bottom: { style: 0, colorArgb: 0xff000000 },
      diagonal: { style: 0, colorArgb: 0xff000000 },
      diagonalUp: false,
      diagonalDown: false,
    }),
    getNumFmtCode: () => null,
  }) as unknown as WorkbookHandle;
