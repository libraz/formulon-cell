import { FormulonError, StatusCode } from '@libraz/formulon';
import { SheetKind } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { zipStored } from './zip-entries.js';

const enc = (text: string): Uint8Array => new TextEncoder().encode(text);

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MACRO_REL = 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet';

/** A minimal package with a worksheet, a macrosheet and a dialogsheet, in that order. */
export function sheetKindWorkbook(): Uint8Array {
  const entries = new Map<string, Uint8Array>([
    [
      '[Content_Types].xml',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/macrosheets/sheet2.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/><Override PartName="/xl/dialogsheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"/></Types>`,
      ),
    ],
    [
      '_rels/.rels',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      ),
    ],
    [
      'xl/workbook.xml',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Macro" sheetId="2" r:id="rId2"/><sheet name="Dialog" sheetId="3" r:id="rId3"/></sheets></workbook>`,
      ),
    ],
    [
      'xl/_rels/workbook.xml.rels',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${MACRO_REL}" Target="macrosheets/sheet2.xml"/><Relationship Id="rId3" Type="${REL}/dialogsheet" Target="dialogsheets/sheet3.xml"/></Relationships>`,
      ),
    ],
    [
      'xl/worksheets/sheet1.xml',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`,
      ),
    ],
    [
      'xl/macrosheets/sheet2.xml',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><macrosheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></macrosheet>`,
      ),
    ],
    [
      'xl/dialogsheets/sheet3.xml',
      enc(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dialogsheet xmlns="${NS}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></dialogsheet>`,
      ),
    ],
  ]);
  return zipStored(entries);
}

const KIND_NAMES: Record<number, string> = {
  [SheetKind.Chartsheet]: 'chartsheet',
  [SheetKind.Dialogsheet]: 'dialogsheet',
  [SheetKind.Macrosheet]: 'macrosheet',
  [SheetKind.Other]: 'sheet',
};

/** Calls that answer on every kind. */
const ANY_KIND = new Set(['kind', 'getVisibility', 'setVisibility']);
/** Calls a chartsheet answers from its own part. */
const CHARTSHEET = new Set([
  'getView',
  'updateView',
  'getTabColor',
  'setTabColor',
  'getPageSetup',
  'updatePageSetup',
  'getPageMargins',
  'updatePageMargins',
  'getHeaderFooter',
  'updateHeaderFooter',
  'getChart',
]);

const refusal = (kind: SheetKind): FormulonError =>
  new FormulonError(
    StatusCode.InvalidArgument,
    `sheet_index is a ${KIND_NAMES[kind] ?? 'sheet'}, but this call needs a worksheet`,
    'sheet_index=0',
  );

/** Make `handle`'s engine refuse, as the engine's kind guard does, every per-sheet call a
 *  sheet of another kind does not answer. Applies to every sheet fetched afterwards. */
export function refuseNonWorksheetCalls(handle: WorkbookHandle): void {
  const raw = (handle as unknown as { wb: { sheet(index: number): { kind: SheetKind } } }).wb;
  const sheetOf = raw.sheet.bind(raw);
  raw.sheet = (index: number) => {
    const sheet = sheetOf(index);
    const kind = sheet.kind;
    if (kind === SheetKind.Worksheet) return sheet;
    return new Proxy(sheet, {
      get(target, prop) {
        const value = Reflect.get(target, prop, target);
        if (typeof value !== 'function' || typeof prop !== 'string') return value;
        const answered =
          ANY_KIND.has(prop) || (kind === SheetKind.Chartsheet && CHARTSHEET.has(prop));
        if (!answered) {
          return () => {
            throw refusal(kind);
          };
        }
        return value.bind(target);
      },
    });
  };
}
