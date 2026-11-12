import type { Color } from '@libraz/formulon';
import { isSheetProtected } from '../../commands/protection.js';
import { activeSparklineGroup } from '../../commands/sparkline-design.js';
import {
  applySparklineStylePresetAction,
  SPARKLINE_STYLE_PRESETS,
  sparklineStylePresetIndex,
} from '../../commands/sparkline-presets.js';
import { cssOfColorSpec } from '../../engine/theme-color.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { ToolbarLang, ToolbarText } from '../ribbon-model.js';
import {
  SPARKLINE_STYLE_SLOTS,
  sparklineStyleLabel,
  sparklineStyleThumbnail,
} from './menus/sparkline.js';
import { createQuickStyleStrip } from './quick-style-strip.js';
import { SPARKLINE_COMMANDS } from './sparkline-tabs.js';

// The desktop strip leaves these presets (Accent 4-6 lighter 40% and Dark #1-#4) to the full
// gallery behind the chevron.
const FULL_GALLERY_ONLY: ReadonlySet<number> = new Set([21, 22, 23, 24, 25, 26, 27]);

export interface SparklineStyleGalleryOptions {
  lang: ToolbarLang;
  text: ToolbarText['sparkline'];
  getInst: () => SpreadsheetInstance | null;
  /** The ribbon's Style button; it opens the full gallery. */
  arrow: HTMLButtonElement;
  menu: HTMLElement | null;
  /** The Mac profile's strip: no page arrows, scrolled to the current style. */
  mac: boolean;
}

/** Sparkline > Style: the strip's 29 presets a row at a time, with the Style button as the more
 *  arrow onto all 36. A pick gives the selected groups the preset's colours. */
export const createSparklineStyleGallery = (opts: SparklineStyleGalleryOptions): HTMLDivElement => {
  const { getInst, text: t } = opts;
  const wrap = createQuickStyleStrip({
    lang: opts.lang,
    classPrefix: 'fc-tb__sparkline-quick',
    kind: 'sparkline',
    mac: opts.mac,
    chipKey: 'sparklineQuickStyle',
    entries: () => {
      const inst = getInst();
      if (!inst) return [];
      const css = (color: Color): string => cssOfColorSpec(inst.workbook, color) ?? '';
      return SPARKLINE_STYLE_SLOTS.flatMap((slot, index) => {
        const preset = SPARKLINE_STYLE_PRESETS[index];
        return preset && !FULL_GALLERY_ONLY.has(index)
          ? [
              {
                id: slot.id,
                label: sparklineStyleLabel(t, index),
                thumbnail: () => sparklineStyleThumbnail(preset, css),
              },
            ]
          : [];
      });
    },
    current: () => {
      const inst = getInst();
      const group = inst ? activeSparklineGroup(inst.store.getState()) : null;
      if (!group) return null;
      return SPARKLINE_STYLE_SLOTS[sparklineStylePresetIndex(group.settings)]?.id ?? '';
    },
    blocked: () => {
      const inst = getInst();
      if (!inst) return { reason: null };
      const state = inst.store.getState();
      if (isSheetProtected(state, state.selection.active.sheet)) {
        return { reason: inst.i18n.strings.mac.sparkline.protectedLocation };
      }
      return activeSparklineGroup(state) ? null : { reason: null };
    },
    apply: (id) => {
      const inst = getInst();
      const index = SPARKLINE_STYLE_SLOTS.findIndex((slot) => slot.id === id);
      if (!inst || index < 0) return;
      applySparklineStylePresetAction(
        {
          store: inst.store,
          history: inst.history,
          origin: 'ribbon',
          commandId: SPARKLINE_COMMANDS.styleGallery,
        },
        index,
      );
      inst.host.focus();
    },
    more: { arrow: opts.arrow, menu: opts.menu },
  });
  opts.arrow.setAttribute('aria-label', t.moreStyles);
  wrap.dataset.ribbonGallery = SPARKLINE_COMMANDS.styleGallery;
  return wrap;
};
