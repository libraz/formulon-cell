export type PasteWhat =
  | 'all'
  | 'all-using-source-theme'
  | 'all-except-borders'
  | 'all-merging-conditional-formats'
  | 'values'
  | 'formulas'
  | 'formats'
  | 'comments'
  | 'validation'
  | 'column-widths'
  | 'formulas-and-numfmt'
  | 'values-and-numfmt';

/** What a paste writes: a Paste Special kind, or the references Paste Link
 *  writes in place of the copied cells. */
export type PasteKind = PasteWhat | 'link';

/** The kinds that paste every layer. A paste within one workbook shares its
 *  theme, so the source-theme variant pastes exactly what All does. */
export const pastesAll = (what: PasteKind): boolean =>
  what === 'all' ||
  what === 'all-using-source-theme' ||
  what === 'all-except-borders' ||
  what === 'all-merging-conditional-formats';
export const wantsValues = (what: PasteKind): boolean =>
  pastesAll(what) || what === 'values' || what === 'values-and-numfmt';
export const wantsFormulas = (what: PasteKind): boolean =>
  pastesAll(what) || what === 'formulas' || what === 'formulas-and-numfmt';
export const wantsFormats = (what: PasteKind): boolean => pastesAll(what) || what === 'formats';
export const wantsNumFmt = (what: PasteKind): boolean =>
  wantsFormats(what) || what === 'values-and-numfmt' || what === 'formulas-and-numfmt';
export const carriesComments = (what: PasteKind): boolean => pastesAll(what) || what === 'comments';
