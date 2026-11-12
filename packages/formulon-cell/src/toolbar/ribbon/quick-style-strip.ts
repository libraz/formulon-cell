import {
  createInteractionButton,
  createInteractionChipButton,
} from '../../interact/chip-button.js';
import { projectDisabledState } from '../menu-a11y.js';
import type { ToolbarLang } from '../ribbon-model.js';

/** One style the strip offers. */
export interface QuickStyleEntry {
  /** What `apply` receives and `current` reports; empty for no style. */
  readonly id: string;
  readonly label: string;
  readonly thumbnail: () => Element;
}

/** The ribbon strips built on {@link createQuickStyleStrip}. */
export type QuickStyleStripKind = 'table' | 'pivot' | 'sparkline';

/** A strip's chip size and pitch in CSS px, and how many chips it shows. */
export interface QuickStyleStripGeometry {
  readonly chipWidth: number;
  readonly chipHeight: number;
  /** Left edge to left edge; below the chip width, neighbours overlap. */
  readonly step: number;
  readonly visible: number;
}

// Mac measured: seven chips each; table 70 x 54 and PivotTable 73 x 56 stepping 71 in a 497 pt
// strip, sparkline 80 x 56 stepping 78. The default profile's strips are unmeasured.
const GEOMETRY: Record<
  QuickStyleStripKind,
  { readonly mac: QuickStyleStripGeometry; readonly desktop: QuickStyleStripGeometry }
> = {
  table: {
    mac: { chipWidth: 70, chipHeight: 54, step: 71, visible: 7 },
    desktop: { chipWidth: 73, chipHeight: 56, step: 75, visible: 7 },
  },
  pivot: {
    mac: { chipWidth: 73, chipHeight: 56, step: 71, visible: 7 },
    desktop: { chipWidth: 73, chipHeight: 56, step: 75, visible: 6 },
  },
  sparkline: {
    mac: { chipWidth: 80, chipHeight: 56, step: 78, visible: 7 },
    desktop: { chipWidth: 80, chipHeight: 56, step: 78, visible: 7 },
  },
};

export const quickStyleStripGeometry = (
  kind: QuickStyleStripKind,
  mac: boolean,
): QuickStyleStripGeometry => GEOMETRY[kind][mac ? 'mac' : 'desktop'];

export interface QuickStyleStripOptions {
  readonly lang: ToolbarLang;
  readonly kind: QuickStyleStripKind;
  /** The Mac strip has no page arrows and scrolls to keep the current style in view. */
  readonly mac: boolean;
  /** Class prefix of the wrapper and its parts: `<prefix>-gallery`, `-strip`, `-chip`, `-prev`,
   *  `-next`. */
  readonly classPrefix: string;
  /** Dataset key each chip carries its entry id under. */
  readonly chipKey: string;
  readonly entries: () => readonly QuickStyleEntry[];
  /** The id of the style in use; null when nothing is selected that takes a style. */
  readonly current: () => string | null;
  readonly apply: (id: string) => void;
  /** Non-null while the chips refuse a click, with the reason to show, if any. */
  readonly blocked?: () => { readonly reason: string | null } | null;
  /** The ribbon's gallery button and its full menu, placed after the strip as its more arrow. */
  readonly more?: { readonly arrow: HTMLButtonElement; readonly menu: HTMLElement | null };
}

const ARROW_LABELS = {
  previous: { ja: '前のクイック スタイル ギャラリー', en: 'Previous Quick Styles gallery' },
  next: { ja: '次のクイック スタイル ギャラリー', en: 'Next Quick Styles gallery' },
} as const;

/** An in-ribbon style gallery: a row of style chips, refreshed by an `fc-ribbon-control-refresh`
 *  event on the wrapper. Whenever the current style changes the row follows it: paging to its page
 *  between arrows, or on the Mac scrolling it to the middle (the fourth of seven). */
export const createQuickStyleStrip = (opts: QuickStyleStripOptions): HTMLDivElement => {
  const { classPrefix } = opts;
  const geometry = quickStyleStripGeometry(opts.kind, opts.mac);
  const pageSize = geometry.visible;
  const arrowLabel = (key: keyof typeof ARROW_LABELS): string =>
    opts.lang === 'ja' ? ARROW_LABELS[key].ja : ARROW_LABELS[key].en;
  const wrap = document.createElement('div');
  wrap.className = `${classPrefix}-gallery`;
  wrap.style.setProperty('--fc-strip-chip-w', `${geometry.chipWidth}px`);
  wrap.style.setProperty('--fc-strip-chip-h', `${geometry.chipHeight}px`);
  wrap.style.setProperty('--fc-strip-step', `${geometry.step}px`);
  wrap.style.setProperty('--fc-strip-visible', String(geometry.visible));
  const strip = document.createElement('div');
  strip.className = `${classPrefix}-strip`;
  strip.setAttribute('role', 'group');
  const prev = createInteractionButton({
    className: `${classPrefix}-prev`,
    ariaLabel: arrowLabel('previous'),
  });
  const next = createInteractionButton({
    className: `${classPrefix}-next`,
    ariaLabel: arrowLabel('next'),
  });
  if (!opts.mac) wrap.append(prev, strip, next);
  else wrap.append(strip);
  if (opts.more) {
    wrap.appendChild(opts.more.arrow);
    if (opts.more.menu) wrap.appendChild(opts.more.menu);
  }

  // The first chip shown; a page start between arrows.
  let start = 0;
  let shownCurrent: string | null | undefined;
  const lastStart = (entries: readonly QuickStyleEntry[]): number =>
    opts.mac
      ? Math.max(0, entries.length - pageSize)
      : Math.max(0, Math.ceil(entries.length / pageSize) - 1) * pageSize;

  const render = (): void => {
    const entries = opts.entries();
    const current = opts.current();
    const blocked = opts.blocked?.() ?? null;
    if (current !== shownCurrent) {
      shownCurrent = current;
      const at = entries.findIndex((entry) => entry.id === current);
      if (at >= 0) {
        start = opts.mac
          ? Math.max(0, at - Math.floor(pageSize / 2))
          : Math.floor(at / pageSize) * pageSize;
      }
    }
    start = Math.min(start, lastStart(entries));
    strip.replaceChildren(
      ...entries.slice(start, start + pageSize).map((entry) => {
        const chip = createInteractionChipButton({
          className: `${classPrefix}-chip`,
          label: '',
          dataset: { [opts.chipKey]: entry.id },
          pressed: current === entry.id,
        });
        chip.setAttribute('aria-label', entry.label);
        chip.title = entry.label;
        chip.appendChild(entry.thumbnail());
        if (blocked) projectDisabledState(chip, true, blocked.reason, { titlePrefix: entry.label });
        chip.addEventListener('click', () => {
          opts.apply(entry.id);
          render();
        });
        return chip;
      }),
    );
    projectDisabledState(prev, start === 0, null);
    projectDisabledState(next, start === lastStart(entries), null);
  };
  wrap.addEventListener('fc-ribbon-control-refresh', () => render());
  prev.addEventListener('click', () => {
    start = Math.max(0, start - pageSize);
    render();
  });
  next.addEventListener('click', () => {
    start = Math.min(lastStart(opts.entries()), start + pageSize);
    render();
  });
  render();
  return wrap;
};
