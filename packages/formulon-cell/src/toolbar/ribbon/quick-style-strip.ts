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

export interface QuickStyleStripOptions {
  readonly lang: ToolbarLang;
  /** Class prefix of the wrapper and its parts: `<prefix>-gallery`, `-strip`, `-chip`, `-prev`,
   *  `-next`. */
  readonly classPrefix: string;
  /** Dataset key each chip carries its entry id under. */
  readonly chipKey: string;
  readonly pageSize: number;
  /** Previous and next page arrows; the Mac strip has none and shows only the current page. */
  readonly arrows: boolean;
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

/** An in-ribbon style gallery: a page of style chips, refreshed by an `fc-ribbon-control-refresh`
 *  event on the wrapper, that turns to the page of the current style whenever it changes. */
export const createQuickStyleStrip = (opts: QuickStyleStripOptions): HTMLDivElement => {
  const { classPrefix, pageSize } = opts;
  const arrowLabel = (key: keyof typeof ARROW_LABELS): string =>
    opts.lang === 'ja' ? ARROW_LABELS[key].ja : ARROW_LABELS[key].en;
  const wrap = document.createElement('div');
  wrap.className = `${classPrefix}-gallery`;
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
  if (opts.arrows) wrap.append(prev, strip, next);
  else wrap.append(strip);
  if (opts.more) {
    wrap.appendChild(opts.more.arrow);
    if (opts.more.menu) wrap.appendChild(opts.more.menu);
  }

  let page = 0;
  let shownCurrent: string | null | undefined;
  const lastPage = (entries: readonly QuickStyleEntry[]): number =>
    Math.max(0, Math.ceil(entries.length / pageSize) - 1);

  const render = (): void => {
    const entries = opts.entries();
    const current = opts.current();
    const blocked = opts.blocked?.() ?? null;
    if (current !== shownCurrent) {
      shownCurrent = current;
      const at = entries.findIndex((entry) => entry.id === current);
      if (at >= 0) page = Math.floor(at / pageSize);
    }
    page = Math.min(page, lastPage(entries));
    strip.replaceChildren(
      ...entries.slice(page * pageSize, page * pageSize + pageSize).map((entry) => {
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
    projectDisabledState(prev, page === 0, null);
    projectDisabledState(next, page === lastPage(entries), null);
  };
  wrap.addEventListener('fc-ribbon-control-refresh', () => render());
  prev.addEventListener('click', () => {
    page = Math.max(0, page - 1);
    render();
  });
  next.addEventListener('click', () => {
    page = Math.min(lastPage(opts.entries()), page + 1);
    render();
  });
  render();
  return wrap;
};
