import {
  type CellRect,
  type Workbook as EngineWorkbook,
  type WorkbookOoxml as EngineWorkbookOoxml,
  FontScheme,
} from '@libraz/formulon';
import { MAX_COL, MAX_ROW, parseA1Atom, rangeOfRect } from '../../engine/address.js';
import { parseA1RangeText } from '../../engine/range-resolver.js';
import type { JournalChoice, JournalCodec, JournalFamily, JournalScope, Restore } from './index.js';

const sheetOf = (scope: JournalScope) => scope.engine.sheet(scope.sheet as number);

const WHOLE_SHEET: CellRect = { firstRow: 0, firstCol: 0, lastRow: MAX_ROW, lastCol: MAX_COL };

const rectKey = (r: CellRect): string => `${r.firstRow}:${r.firstCol}:${r.lastRow}:${r.lastCol}`;

/** The rectangle a CellRect or A1 range argument names, or the whole sheet when it does not parse. */
const rectOfArg = (arg: unknown): CellRect => {
  if (typeof arg === 'object' && arg !== null) return arg as CellRect;
  const parsed = typeof arg === 'string' ? parseA1RangeText(arg) : null;
  return parsed
    ? { firstRow: parsed.r0, firstCol: parsed.c0, lastRow: parsed.r1, lastCol: parsed.c1 }
    : WHOLE_SHEET;
};

/** The cell a `(row, col)` or A1 argument pair names. */
const cellOfArgs = (args: readonly unknown[]): { row: number; col: number } | null =>
  typeof args[0] === 'number'
    ? { row: args[0], col: args[1] as number }
    : parseA1Atom(String(args[0]));

/** The cells a facade write on the sheet touches, from its range or its arguments. */
const touchedRect = (scope: JournalScope): CellRect => {
  if (scope.range) return scope.range;
  const { name, args } = scope;
  if (name === 'setCellXfIndex' || name === 'resetCellXfIndex') {
    const cell = cellOfArgs(args);
    return cell
      ? { firstRow: cell.row, firstCol: cell.col, lastRow: cell.row, lastCol: cell.col }
      : WHOLE_SHEET;
  }
  if (name.endsWith('RowXfIndex') || name.endsWith('ColumnXfIndex')) {
    const byRow = name.endsWith('RowXfIndex');
    const span = rectOfArg(typeof args[0] === 'number' ? undefined : args[0]);
    const first = typeof args[0] === 'number' ? args[0] : byRow ? span.firstRow : span.firstCol;
    const last =
      typeof args[0] === 'number' ? (args[1] as number) : byRow ? span.lastRow : span.lastCol;
    return byRow
      ? { firstRow: first, firstCol: 0, lastRow: last, lastCol: MAX_COL }
      : { firstRow: 0, firstCol: first, lastRow: MAX_ROW, lastCol: last };
  }
  if (name === 'setRangeXfIndex' || name === 'remapRangeXfs') return rectOfArg(args[0]);
  return WHOLE_SHEET;
};

/** Captures the XFs of `rects`; restoring writes each snapshot back. `null` without a range snapshot. */
const captureXfs = (
  scope: JournalScope,
  sheet: number,
  rects: readonly CellRect[],
): Restore | null => {
  const snapshots: Uint8Array[] = [];
  for (const rect of rects) {
    const bytes = scope.handle.snapshotRangeStyle(rangeOfRect(sheet, rect));
    if (!bytes) return null;
    snapshots.push(bytes);
  }
  return () => {
    for (const bytes of snapshots) scope.handle.restoreRangeStyle(sheet, bytes);
  };
};

const XF_RANGE: JournalCodec = {
  id: 'xf-range',
  scopeKey: (scope) => `${scope.sheet}:${rectKey(touchedRect(scope))}`,
  capture: (scope) => captureXfs(scope, scope.sheet as number, [touchedRect(scope)]),
};

/** The ranges of the first capture, so the capture after the write covers them too. */
const linkRanges = new WeakMap<JournalScope, Map<string, CellRect>>();

const HYPERLINKS: JournalCodec = {
  id: 'hyperlinks',
  scopeKey: (scope) => `${scope.sheet}`,
  capture: (scope) => {
    const sheet = scope.sheet as number;
    const engineSheet = sheetOf(scope);
    const links = engineSheet.listHyperlinks().map((link) => ({ ...link }));
    const ranges = linkRanges.get(scope) ?? new Map<string, CellRect>();
    linkRanges.set(scope, ranges);
    const rects = links.map((link) => link.range);
    if (scope.name === 'setHyperlink') {
      rects.push(rectOfArg((scope.args[0] as { range: CellRect | string }).range));
    }
    for (const rect of rects) ranges.set(rectKey(rect), rect);
    const restoreXfs = captureXfs(scope, sheet, [...ranges.values()]);
    if (!restoreXfs) return null;
    return () => {
      engineSheet.clearHyperlinks();
      links.forEach((link, i) => {
        engineSheet.insertHyperlinkAt(i, link);
      });
      restoreXfs();
    };
  },
};

const PHONETIC: JournalCodec = {
  id: 'phonetic',
  scopeKey: (scope) => {
    const cell = cellOfArgs(scope.args);
    return `${scope.sheet}:${cell?.row}:${cell?.col}`;
  },
  capture: (scope) => {
    const cell = cellOfArgs(scope.args);
    if (!cell) return null;
    const { row, col } = cell;
    const sheet = sheetOf(scope);
    const text = sheet.getCellPhonetic(row, col);
    const runs = sheet.getCellPhoneticRuns(row, col);
    const properties = sheet.getCellPhoneticProperties(row, col);
    const show = sheet.getCellShowPhonetic(row, col);
    return () => {
      sheet.setCellPhonetic(row, col, text);
      sheet.setCellPhoneticRuns(row, col, runs.length > 0 ? runs : null);
      sheet.setCellPhoneticProperties(row, col, properties.stated ? properties : null);
      sheet.setCellShowPhonetic(row, col, show);
    };
  },
};

const FMT_DEFAULTS: JournalCodec = {
  id: 'format-defaults',
  scopeKey: (scope) => `${scope.sheet}`,
  capture: (scope) => {
    const sheet = sheetOf(scope);
    const {
      defaultColumnWidthChars,
      defaultRowHeightPt,
      baseColumnWidthChars,
      customHeight,
      zeroHeight,
    } = sheet.getFormatDefaults();
    return () =>
      sheet.updateFormatDefaults({
        defaultColumnWidthChars,
        defaultRowHeightPt,
        baseColumnWidthChars,
        customHeight,
        zeroHeight,
      });
  },
};

const argbHex = (argb: number): string => `#${(argb >>> 0).toString(16).padStart(8, '0')}`;

const WB_STYLES: JournalCodec = {
  id: 'workbook-styles',
  scopeKey: () => 'workbook',
  capture: (scope) => {
    const wb = scope.engine;
    const styles = wb.listCellStyles();
    const defaultFont = wb.getDefaultFont();
    const colors = wb.getThemeColors().colors.map(argbHex);
    const fonts = wb.getThemeFonts();
    const scripts = [FontScheme.Major, FontScheme.Minor].flatMap((scheme) =>
      wb.listThemeScriptFonts(scheme),
    );
    return () => {
      const kept = new Set(styles.map((style) => style.name));
      for (const style of wb.listCellStyles()) {
        if (!kept.has(style.name)) wb.removeCellStyle(style.name);
      }
      for (const style of styles) wb.setCellStyle(style);
      wb.setDefaultFont(defaultFont);
      wb.setThemeColors({ colors });
      wb.setThemeFonts(fonts);
      for (const script of scripts) wb.setThemeScriptFont(script);
    };
  },
};

type OoxmlRecord = (ooxml: EngineWorkbookOoxml, index: number) => Restore;

const OOXML_RECORDS: Readonly<Record<string, OoxmlRecord>> = {
  setCellXf: (o, i) => {
    const record = o.getCellXf(i);
    return () => o.setCellXf(i, record);
  },
  setCellStyleXf: (o, i) => {
    const record = o.getCellStyleXf(i);
    return () => o.setCellStyleXf(i, record);
  },
  setFont: (o, i) => {
    const record = o.getFont(i);
    return () => o.setFont(i, record);
  },
  setFill: (o, i) => {
    const record = o.getFill(i);
    return () => o.setFill(i, record);
  },
  setBorder: (o, i) => {
    const record = o.getBorder(i);
    return () => o.setBorder(i, record);
  },
  setDxf: (o, i) => {
    const record = o.getDxf(i);
    return () => o.setDxf(i, record);
  },
};

const OOXML_REC: JournalCodec = {
  id: 'ooxml-record',
  scopeKey: (scope) => `${scope.name}:${scope.args[0]}`,
  capture: (scope) => {
    const read = OOXML_RECORDS[scope.name];
    return read ? read((scope.engine as EngineWorkbook).ooxml, scope.args[0] as number) : null;
  },
};

const members = (codec: JournalChoice, names: readonly string[]) =>
  Object.fromEntries(names.map((name) => [name, codec]));

export const FORMAT_JOURNAL: JournalFamily = {
  members: {
    sheet: {
      ...members(HYPERLINKS, ['setHyperlink', 'removeHyperlink', 'clearHyperlinks']),
      ...members(PHONETIC, [
        'setCellPhonetic',
        'setCellPhoneticRuns',
        'setCellPhoneticProperties',
        'setCellShowPhonetic',
      ]),
      updateFormatDefaults: FMT_DEFAULTS,
    },
    'sheet.ooxml': members(XF_RANGE, [
      'setCellXfIndex',
      'resetCellXfIndex',
      'setRowXfIndex',
      'resetRowXfIndex',
      'setColumnXfIndex',
      'resetColumnXfIndex',
      'setRangeXfIndex',
      'remapRangeXfs',
      'restoreRangeXfs',
    ]),
    workbook: members(WB_STYLES, [
      'setCellStyle',
      'removeCellStyle',
      'setDefaultFont',
      'setThemeColors',
      'setThemeFonts',
      'setThemeScriptFont',
      'resetTheme',
    ]),
    'workbook.ooxml': {
      ...members(OOXML_REC, Object.keys(OOXML_RECORDS)),
      ...members('none', [
        'addCellXf',
        'addCellStyleXf',
        'addFont',
        'addFill',
        'addBorder',
        'addNumberFormat',
        'addDxf',
      ]),
    },
  },
  kinds: { 'format/range': XF_RANGE },
};
