import type { WorkbookHandle } from './workbook-handle.js';

/** The two workbook theme presets exposed by the Page Layout and Mac ribbon. */
export type WorkbookThemePresetId = 'office' | 'gray';

export interface WorkbookThemePreset {
  readonly colors: readonly number[];
  readonly fonts: {
    readonly majorLatin: string;
    readonly majorEastAsian: string;
    readonly minorLatin: string;
    readonly minorEastAsian: string;
  };
}

/** clrScheme order: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink. */
export const WORKBOOK_THEME_PRESETS: Readonly<Record<WorkbookThemePresetId, WorkbookThemePreset>> =
  {
    office: {
      colors: [
        0xff000000, 0xffffffff, 0xff1f497d, 0xffeeece1, 0xff4f81bd, 0xffc0504d, 0xff9bbb59,
        0xff8064a2, 0xff4bacc6, 0xfff79646, 0xff0000ff, 0xff800080,
      ],
      fonts: {
        majorLatin: 'Calibri Light',
        majorEastAsian: '游ゴシック Light',
        minorLatin: 'Calibri',
        minorEastAsian: '游ゴシック',
      },
    },
    gray: {
      colors: [
        0xff000000, 0xffffffff, 0xff262626, 0xfff2f2f2, 0xff7f7f7f, 0xffa6a6a6, 0xff595959,
        0xffbfbfbf, 0xffd9d9d9, 0xff404040, 0xff0563c1, 0xff954f72,
      ],
      fonts: {
        majorLatin: 'Arial',
        majorEastAsian: 'ＭＳ Ｐゴシック',
        minorLatin: 'Arial',
        minorEastAsian: 'ＭＳ Ｐゴシック',
      },
    },
  };

/** Whether the complete theme currently matches a preset. */
export function workbookThemeMatchesPreset(
  workbook: WorkbookHandle,
  presetId: WorkbookThemePresetId,
): boolean {
  const current = workbook.getTheme();
  const preset = WORKBOOK_THEME_PRESETS[presetId];
  return (
    current?.colors.join(',') === preset.colors.join(',') &&
    current?.fonts.majorLatin === preset.fonts.majorLatin &&
    current?.fonts.majorEastAsian === preset.fonts.majorEastAsian &&
    current?.fonts.minorLatin === preset.fonts.minorLatin &&
    current?.fonts.minorEastAsian === preset.fonts.minorEastAsian
  );
}
