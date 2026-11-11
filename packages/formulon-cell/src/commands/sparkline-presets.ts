import type { Color } from '@libraz/formulon';
import { ColorKind } from '@libraz/formulon';
import type { SparklineGroupSettings } from '../store/store.js';
import type { ObjectActionContext } from './object-actions.js';
import { patchSelectedSparklineGroupsAction } from './sparkline-design.js';

/** The colours a Style preset sets, one per part of the sparkline. */
export interface SparklineStylePreset {
  readonly series: Color;
  readonly negative: Color;
  readonly axis: Color;
  readonly markers: Color;
  readonly first: Color;
  readonly last: Color;
  readonly high: Color;
  readonly low: Color;
}

const PRESET_KEYS: Readonly<Record<keyof SparklineStylePreset, keyof SparklineGroupSettings>> = {
  series: 'colorSeries',
  negative: 'colorNegative',
  axis: 'colorAxis',
  markers: 'colorMarkers',
  first: 'colorFirst',
  last: 'colorLast',
  high: 'colorHigh',
  low: 'colorLow',
};

/** A theme colour; the tint is the saved file's exact string, so it stays theme-relative. */
const theme = (index: number, tint = '0'): Color => ({
  kind: ColorKind.Theme,
  argb: 0,
  theme: index,
  tint: Number(tint),
  indexed: 0,
});

const rgb = (argb: number): Color => ({
  kind: ColorKind.Rgb,
  argb,
  theme: 0,
  tint: 0,
  indexed: 0,
});

const preset = (
  series: Color,
  negative: Color,
  axis: Color,
  markers: Color,
  first: Color,
  last: Color,
  high: Color,
  low: Color,
): SparklineStylePreset => ({ series, negative, axis, markers, first, last, high, low });

/** The Style gallery's presets in gallery order (six rows of six), as the desktop app saves them
 *  under the default Office theme: four accent families, then Dark #1-#6 and Colorful #1-#6. */
export const SPARKLINE_STYLE_PRESETS: readonly SparklineStylePreset[] = [
  preset(
    theme(4, '-0.499984740745262'),
    theme(5),
    rgb(0xff000000),
    theme(4, '-0.499984740745262'),
    theme(4, '0.39997558519241921'),
    theme(4, '0.39997558519241921'),
    theme(4),
    theme(4),
  ),
  preset(
    theme(5, '-0.499984740745262'),
    theme(6),
    rgb(0xff000000),
    theme(5, '-0.499984740745262'),
    theme(5, '0.39997558519241921'),
    theme(5, '0.39997558519241921'),
    theme(5),
    theme(5),
  ),
  preset(
    theme(6, '-0.499984740745262'),
    theme(7),
    rgb(0xff000000),
    theme(6, '-0.499984740745262'),
    theme(6, '0.39997558519241921'),
    theme(6, '0.39997558519241921'),
    theme(6),
    theme(6),
  ),
  preset(
    theme(7, '-0.499984740745262'),
    theme(8),
    rgb(0xff000000),
    theme(7, '-0.499984740745262'),
    theme(7, '0.39997558519241921'),
    theme(7, '0.39997558519241921'),
    theme(7),
    theme(7),
  ),
  preset(
    theme(8, '-0.499984740745262'),
    theme(9),
    rgb(0xff000000),
    theme(8, '-0.499984740745262'),
    theme(8, '0.39997558519241921'),
    theme(8, '0.39997558519241921'),
    theme(8),
    theme(8),
  ),
  preset(
    theme(9, '-0.499984740745262'),
    theme(4),
    rgb(0xff000000),
    theme(9, '-0.499984740745262'),
    theme(9, '0.39997558519241921'),
    theme(9, '0.39997558519241921'),
    theme(9),
    theme(9),
  ),
  preset(
    theme(4, '-0.249977111117893'),
    theme(5),
    rgb(0xff000000),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
  ),
  preset(
    theme(5, '-0.249977111117893'),
    theme(6),
    rgb(0xff000000),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
  ),
  preset(
    theme(6, '-0.249977111117893'),
    theme(7),
    rgb(0xff000000),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
  ),
  preset(
    theme(7, '-0.249977111117893'),
    theme(8),
    rgb(0xff000000),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
  ),
  preset(
    theme(8, '-0.249977111117893'),
    theme(9),
    rgb(0xff000000),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
  ),
  preset(
    theme(9, '-0.249977111117893'),
    theme(4),
    rgb(0xff000000),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
  ),
  preset(
    theme(4),
    theme(5),
    rgb(0xff000000),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
  ),
  preset(
    theme(5),
    theme(6),
    rgb(0xff000000),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
  ),
  preset(
    theme(6),
    theme(7),
    rgb(0xff000000),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
  ),
  preset(
    theme(7),
    theme(8),
    rgb(0xff000000),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
  ),
  preset(
    theme(8),
    theme(9),
    rgb(0xff000000),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
  ),
  preset(
    theme(9),
    theme(4),
    rgb(0xff000000),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
  ),
  preset(
    theme(4, '0.39997558519241921'),
    theme(0, '-0.499984740745262'),
    rgb(0xff000000),
    theme(4, '0.79998168889431442'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.249977111117893'),
    theme(4, '-0.499984740745262'),
    theme(4, '-0.499984740745262'),
  ),
  preset(
    theme(5, '0.39997558519241921'),
    theme(0, '-0.499984740745262'),
    rgb(0xff000000),
    theme(5, '0.79998168889431442'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.249977111117893'),
    theme(5, '-0.499984740745262'),
    theme(5, '-0.499984740745262'),
  ),
  preset(
    theme(6, '0.39997558519241921'),
    theme(0, '-0.499984740745262'),
    rgb(0xff000000),
    theme(6, '0.79998168889431442'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.249977111117893'),
    theme(6, '-0.499984740745262'),
    theme(6, '-0.499984740745262'),
  ),
  preset(
    theme(7, '0.39997558519241921'),
    theme(0, '-0.499984740745262'),
    rgb(0xff000000),
    theme(7, '0.79998168889431442'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.249977111117893'),
    theme(7, '-0.499984740745262'),
    theme(7, '-0.499984740745262'),
  ),
  preset(
    theme(8, '0.39997558519241921'),
    theme(0, '-0.499984740745262'),
    rgb(0xff000000),
    theme(8, '0.79998168889431442'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.249977111117893'),
    theme(8, '-0.499984740745262'),
    theme(8, '-0.499984740745262'),
  ),
  preset(
    theme(9, '0.39997558519241921'),
    theme(0, '-0.499984740745262'),
    rgb(0xff000000),
    theme(9, '0.79998168889431442'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.249977111117893'),
    theme(9, '-0.499984740745262'),
    theme(9, '-0.499984740745262'),
  ),
  preset(
    theme(3, '0.499984740745262'),
    theme(1, '0.249977111117893'),
    rgb(0xff000000),
    theme(1, '0.249977111117893'),
    theme(1, '0.249977111117893'),
    theme(1, '0.249977111117893'),
    theme(1, '0.249977111117893'),
    theme(1, '0.249977111117893'),
  ),
  preset(
    theme(3, '0.34998626667073579'),
    theme(0, '-0.249977111117893'),
    rgb(0xff000000),
    theme(0, '-0.249977111117893'),
    theme(0, '-0.249977111117893'),
    theme(0, '-0.249977111117893'),
    theme(0, '-0.249977111117893'),
    theme(0, '-0.249977111117893'),
  ),
  preset(
    rgb(0xff323232),
    rgb(0xffd00000),
    rgb(0xff000000),
    rgb(0xffd00000),
    rgb(0xffd00000),
    rgb(0xffd00000),
    rgb(0xffd00000),
    rgb(0xffd00000),
  ),
  preset(
    rgb(0xff000000),
    rgb(0xff0070c0),
    rgb(0xff000000),
    rgb(0xff0070c0),
    rgb(0xff0070c0),
    rgb(0xff0070c0),
    rgb(0xff0070c0),
    rgb(0xff0070c0),
  ),
  preset(
    rgb(0xff376092),
    rgb(0xffd00000),
    rgb(0xff000000),
    rgb(0xffd00000),
    rgb(0xffd00000),
    rgb(0xffd00000),
    rgb(0xffd00000),
    rgb(0xffd00000),
  ),
  preset(
    rgb(0xff0070c0),
    rgb(0xff000000),
    rgb(0xff000000),
    rgb(0xff000000),
    rgb(0xff000000),
    rgb(0xff000000),
    rgb(0xff000000),
    rgb(0xff000000),
  ),
  preset(
    rgb(0xff5f5f5f),
    rgb(0xffffb620),
    rgb(0xff000000),
    rgb(0xffd70077),
    rgb(0xff5687c2),
    rgb(0xff359ceb),
    rgb(0xff56be79),
    rgb(0xffff5055),
  ),
  preset(
    rgb(0xff5687c2),
    rgb(0xffffb620),
    rgb(0xff000000),
    rgb(0xffd70077),
    rgb(0xff777777),
    rgb(0xff359ceb),
    rgb(0xff56be79),
    rgb(0xffff5055),
  ),
  preset(
    rgb(0xffc6efce),
    rgb(0xffffc7ce),
    rgb(0xff000000),
    rgb(0xff8cadd6),
    rgb(0xffffdc47),
    rgb(0xffffeb9c),
    rgb(0xff60d276),
    rgb(0xffff5367),
  ),
  preset(
    rgb(0xff00b050),
    rgb(0xffff0000),
    rgb(0xff000000),
    rgb(0xff0070c0),
    rgb(0xffffc000),
    rgb(0xffffc000),
    rgb(0xff00b050),
    rgb(0xffff0000),
  ),
  preset(theme(3), theme(9), rgb(0xff000000), theme(8), theme(4), theme(5), theme(6), theme(7)),
  preset(theme(3), theme(4), rgb(0xff000000), theme(9), theme(5), theme(6), theme(7), theme(8)),
];

const sameColor = (a: Color | null, b: Color): boolean =>
  a !== null &&
  a.kind === b.kind &&
  (b.kind === ColorKind.Rgb ? a.argb === b.argb : a.theme === b.theme && a.tint === b.tint);

/** The gallery index of the preset whose colours `settings` carry, or -1 for none. */
export const sparklineStylePresetIndex = (settings: SparklineGroupSettings | undefined): number =>
  settings
    ? SPARKLINE_STYLE_PRESETS.findIndex((p) =>
        (Object.keys(PRESET_KEYS) as (keyof SparklineStylePreset)[]).every((part) =>
          sameColor(settings[PRESET_KEYS[part]] as Color | null, p[part]),
        ),
      )
    : -1;

/** Give the selected groups every colour of the preset at gallery `index`. */
export function applySparklineStylePresetAction(ctx: ObjectActionContext, index: number): boolean {
  const picked = SPARKLINE_STYLE_PRESETS[index];
  if (!picked) return false;
  const patch: Partial<SparklineGroupSettings> = {};
  for (const part of Object.keys(PRESET_KEYS) as (keyof SparklineStylePreset)[]) {
    Object.assign(patch, { [PRESET_KEYS[part]]: picked[part] });
  }
  return patchSelectedSparklineGroupsAction(ctx, patch);
}
