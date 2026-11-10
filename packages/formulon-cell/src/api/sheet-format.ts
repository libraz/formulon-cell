import {
  type CellFormatPatch,
  type CellRect,
  type Sheet as EngineSheet,
  type Workbook as EngineWorkbook,
  FormulonError,
  type SheetProtectionInput,
  StatusCode,
} from '@libraz/formulon';
import {
  hashSheetProtectionPassword,
  legacySheetPasswordHash,
  verifySheetProtectionPasswordHash,
} from '../engine/protection-password.js';
import { spreadsheetFailure } from './errors.js';
import type { MemberSpec, SheetContext, WorkbookContext } from './facade-context.js';

/** The password-hash keys of a protection record, which `Sheet.protect` derives from a password. */
type ProtectionHashKey =
  | 'algorithmName'
  | 'hashValue'
  | 'saltValue'
  | 'spinCount'
  | 'legacyPasswordHash';

/** Options of `Sheet.protect`: the `allow*` flags of the protection, and an optional password. */
export interface ProtectOptions extends Omit<SheetProtectionInput, ProtectionHashKey> {
  /** The password `unprotect` will ask for; none by default. */
  readonly password?: string;
}

async function protect(ctx: SheetContext, options: ProtectOptions = {}): Promise<void> {
  const { password, ...flags } = options;
  let hash = {};
  if (password !== undefined) {
    if (typeof password !== 'string') throw new TypeError('protect: password must be a string');
    const derived = await hashSheetProtectionPassword(password);
    if (derived === null) {
      throw new FormulonError(
        StatusCode.NotImplemented,
        'protect: hashing a password needs Web Crypto (crypto.subtle)',
      );
    }
    hash = derived;
  }
  ctx.sheetChange('protection', () => ctx.sheetEngine.setSheetProtection({ ...flags, ...hash }));
}

async function unprotect(ctx: SheetContext, password?: string): Promise<void> {
  const current = ctx.sheetEngine.getSheetProtection();
  if (current === null) return;
  const given = password ?? '';
  const matches =
    current.hashValue !== null && current.saltValue !== null && current.algorithmName !== null
      ? await verifySheetProtectionPasswordHash(given, {
          algorithmName: current.algorithmName,
          hashValue: current.hashValue,
          saltValue: current.saltValue,
          spinCount: current.spinCount,
        })
      : current.legacyPasswordHash === null ||
        current.legacyPasswordHash.toUpperCase() === legacySheetPasswordHash(given);
  if (!matches) {
    throw spreadsheetFailure('protected', 'unprotect: the password does not match');
  }
  ctx.sheetChange('protection', () => ctx.sheetEngine.setSheetProtection(null));
}

/**
 * Cell formats, conditional formats, data validation, protection and print settings: engine
 * `Sheet` members of the same name, plus the password convenience `protect`/`unprotect`.
 */
export const SHEET_FORMAT_MEMBERS = {
  read: [
    'getCellFormat',
    'getFormatDefaults',
    'getCfRule',
    'listCfRules',
    'evaluateCfRange',
    'evaluateCfFormula',
    'getValidation',
    'listValidations',
    'validateValue',
    'getInvalidCells',
    'getSheetProtection',
    'getProtectedRange',
    'listProtectedRanges',
    'getPrintArea',
    'getPrintTitles',
    'getPageSetup',
    'getPageMargins',
    'getPrintOptions',
    'getHeaderFooter',
    'getPagination',
    'listRowBreaks',
    'listColumnBreaks',
  ],
  change: {
    format: ['updateFormatDefaults'],
    cf: ['addCfRule', 'updateCfRule', 'removeCfRule', 'clearCfRules'],
    validation: ['addValidation', 'updateValidation', 'removeValidation', 'clearValidations'],
    protection: [
      'setSheetProtection',
      'updateSheetProtection',
      'addProtectedRange',
      'updateProtectedRange',
      'removeProtectedRange',
    ],
    print: [
      'setPrintArea',
      'setPrintTitles',
      'updatePageSetup',
      'updatePageMargins',
      'updatePrintOptions',
      'updateHeaderFooter',
      'setRowBreak',
      'setColumnBreak',
      'clearBreaks',
    ],
  },
  methods: {
    protect,
    unprotect,
    updateCellFormat: (ctx: SheetContext, range: CellRect | string, patch: CellFormatPatch): void =>
      ctx.sheetChange(
        'format',
        () => ctx.sheetEngine.updateCellFormat(range, patch),
        ctx.rectOf(range),
      ),
  },
} as const satisfies MemberSpec<SheetContext>;

/** Cell styles, the default font, the theme and workbook protection: engine `Workbook` members. */
export const WORKBOOK_FORMAT_MEMBERS = {
  read: [
    'getCellStyle',
    'listCellStyles',
    'getDefaultFont',
    'getThemeColors',
    'getThemeFonts',
    'getThemeScriptFont',
    'listThemeScriptFonts',
    'resolveColor',
    'getWorkbookProtection',
  ],
  change: {
    format: [
      'setCellStyle',
      'removeCellStyle',
      'setDefaultFont',
      'setThemeColors',
      'setThemeFonts',
      'setThemeScriptFont',
      'resetTheme',
    ],
    protection: ['setWorkbookProtection'],
  },
} as const satisfies MemberSpec<WorkbookContext>;

type SheetChanges = typeof SHEET_FORMAT_MEMBERS.change;
type SheetFormatMemberName =
  | (typeof SHEET_FORMAT_MEMBERS.read)[number]
  | SheetChanges[keyof SheetChanges][number]
  | 'updateCellFormat';
type WorkbookChanges = typeof WORKBOOK_FORMAT_MEMBERS.change;
type WorkbookFormatMemberName =
  | (typeof WORKBOOK_FORMAT_MEMBERS.read)[number]
  | WorkbookChanges[keyof WorkbookChanges][number];

declare module './workbook.js' {
  interface Sheet extends Pick<EngineSheet, SheetFormatMemberName> {
    /**
     * Protects the sheet with the given `allow*` flags. A `password` is hashed (SHA-512, salted,
     * 100,000 spins, as Excel writes it) before it reaches the workbook; the text itself is not kept.
     *
     * @throws FormulonError `NotImplemented` when a password is given and Web Crypto is unavailable.
     */
    protect(options?: ProtectOptions): Promise<void>;
    /**
     * Removes the sheet's protection. A protected sheet with a password needs that password; an
     * unprotected sheet is left as it is.
     *
     * @throws FormulonError `protected` when the password does not match.
     */
    unprotect(password?: string): Promise<void>;
  }
  interface Workbook extends Pick<EngineWorkbook, WorkbookFormatMemberName> {}
}
