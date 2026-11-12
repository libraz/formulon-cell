/**
 * Width fitting for the ribbon panel.
 *
 * When the visible panel is wider than its viewport, tile groups give up width
 * in stages, each applied from the right edge inward: tiles narrow so labels
 * wrap onto a second line, then large buttons drop to icon-and-label rows
 * stacked three high, and finally those rows drop their labels (the accessible
 * name and tooltip stay). Command order never changes; only density does. A
 * stage is kept on a group only if it made the group narrower. On the Mac
 * profile a panel that still overflows then collapses whole groups, rightmost
 * first, into one named button that opens the group as a popup.
 */

import { onOutsidePress } from '../../interact/outside-press.js';
import { clampPanelToViewport, viewportSize } from '../../interact/overlay-position.js';
import { createRibbonButton } from './button.js';

export const RIBBON_GROUP_NARROW_CLASS = 'fc-tb__ribbon-group--narrow';
export const RIBBON_GROUP_CONDENSED_CLASS = 'fc-tb__ribbon-group--condensed';
export const RIBBON_GROUP_ICON_ONLY_CLASS = 'fc-tb__ribbon-group--icon-only';
export const RIBBON_GROUP_COLLAPSED_CLASS = 'fc-tb__ribbon-group--collapsed';
const GROUP_OPEN_CLASS = 'fc-tb__ribbon-group--open';
const COLLAPSE_BUTTON_CLASS = 'fc-tb__ribbon-group-collapse';
const GROUP_POPUP_PAD = 8;
const GROUP_POPUP_GAP = 2;

const isTileGroup = (group: HTMLElement): boolean =>
  group.classList.contains('fc-tb__ribbon-group--tiles') ||
  group.classList.contains('fc-tb__ribbon-group--mixed');

const rowButtons = (group: HTMLElement): HTMLElement[] => [
  ...group.querySelectorAll<HTMLElement>(':scope > .fc-tb__ribbon-tools > .fc-tb__rb'),
];

// Only groups already laid out as rows, and never one with text-only buttons:
// a bare "75%" beside icon-only neighbours reads as their label.
const isRowGroup = (group: HTMLElement): boolean => {
  const buttons = rowButtons(group);
  const rows =
    group.classList.contains(RIBBON_GROUP_CONDENSED_CLASS) ||
    buttons.some((button) => button.classList.contains('fc-tb__rb--stacked'));
  return rows && !buttons.some((button) => button.classList.contains('fc-tb__rb--mono'));
};

/** Each stage and the groups it may apply to. */
const FIT_STAGES: readonly (readonly [string, (group: HTMLElement) => boolean])[] = [
  [RIBBON_GROUP_NARROW_CLASS, isTileGroup],
  [RIBBON_GROUP_CONDENSED_CLASS, isTileGroup],
  [RIBBON_GROUP_ICON_ONLY_CLASS, isRowGroup],
];

const FIT_CLASSES = FIT_STAGES.map(([stage]) => stage);

const LABEL = 'span:not(.fc-tb__rb-icon):not(.fc-tb__rb-split-chevron)';
/** Labels of a group's full-height tiles. */
const TILE_LABELS = `:scope > .fc-tb__ribbon-tools > .fc-tb__rb:not(.fc-tb__rb--stacked) > ${LABEL}`;
/** Mac Home style and editing tiles: Excel sizes each to the narrowest width
 *  that keeps its label on two lines. */
const TWO_LINE_LABELS = `:scope > .fc-tb__ribbon :is(.fc-tb__ribbon-group--styles, .fc-tb__ribbon-group--editing) .fc-tb__rb--wide:not(.fc-tb__rb--stacked) > ${LABEL}`;

const overflows = (panel: HTMLElement): boolean => panel.scrollWidth > panel.clientWidth + 1;

const clearFit = (groups: Iterable<Element>): void => {
  for (const group of groups) {
    group.classList.remove(...FIT_CLASSES, RIBBON_GROUP_COLLAPSED_CLASS, GROUP_OPEN_CLASS);
    group
      .querySelector(`:scope > .${COLLAPSE_BUTTON_CLASS}`)
      ?.setAttribute('aria-expanded', 'false');
  }
};

/** A group of several commands can collapse; the style gallery group and one-button groups never do. */
const canCollapse = (group: HTMLElement): boolean =>
  !group.classList.contains('fc-tb__ribbon-group--styles') &&
  group.querySelectorAll('[data-ribbon-command]').length > 1;

/** Where focus goes back to for `el`: the button that folds its group while the group is folded
 *  and closed, since `el` is hidden then; otherwise `el` itself. */
export const visibleFocusReturn = (el: HTMLElement): HTMLElement => {
  const group = el.closest<HTMLElement>(`.${RIBBON_GROUP_COLLAPSED_CLASS}`);
  if (!group || group.classList.contains(GROUP_OPEN_CLASS)) return el;
  return group.querySelector<HTMLElement>(`:scope > .${COLLAPSE_BUTTON_CLASS}`) ?? el;
};

/** Document listeners an open popup holds; closing the popup releases them. */
const popupListeners = new WeakMap<HTMLElement, () => void>();

const closeGroupPopup = (group: HTMLElement): void => {
  popupListeners.get(group)?.();
  popupListeners.delete(group);
  group.classList.remove(GROUP_OPEN_CLASS);
  group.querySelector(`:scope > .${COLLAPSE_BUTTON_CLASS}`)?.setAttribute('aria-expanded', 'false');
  const tools = group.querySelector<HTMLElement>(':scope > .fc-tb__ribbon-tools');
  tools?.style.removeProperty('left');
  tools?.style.removeProperty('top');
  tools?.style.removeProperty('max-width');
  tools?.style.removeProperty('max-height');
};

/** Controls that open their own list or menu; a click on them must keep the popup open. */
const OPENS_OWN_LIST = '[aria-haspopup], .fc-tb__rb-dd, .fc-tb__rb-color, .fc-tb__menu';
/** Lists and menus a control inside the popup opens; they may live outside the group. */
const OPENED_LIST = '.fc-tb__menu, .fc-tb__rb-dd__list';
/** What runs a command or picks a value: choosing one dismisses the whole popup chain. */
const LEAF = '[data-ribbon-command], [role="menuitem"], [role="option"]';

/** Whether a click at `target` chose a command or value rather than opening a nested list. */
const choseLeaf = (target: Element): boolean => {
  if (target.closest(OPENED_LIST)) {
    const leaf = target.closest(LEAF);
    return leaf !== null && !leaf.matches('[aria-haspopup]:not([aria-haspopup="false"])');
  }
  return target.closest('[data-ribbon-command]') !== null && !target.closest(OPENS_OWN_LIST);
};

const openGroupPopup = (group: HTMLElement, button: HTMLElement): void => {
  group.classList.add(GROUP_OPEN_CLASS);
  button.setAttribute('aria-expanded', 'true');
  const tools = group.querySelector<HTMLElement>(':scope > .fc-tb__ribbon-tools');
  const rect = button.getBoundingClientRect();
  if (tools) {
    // Cap the popup to the viewport (its tools wrap and scroll), then slide it back inside.
    const viewport = viewportSize(group.ownerDocument);
    const desiredY = rect.bottom + GROUP_POPUP_GAP;
    tools.style.maxWidth = `${Math.max(0, viewport.width - GROUP_POPUP_PAD * 2)}px`;
    tools.style.maxHeight = `${Math.max(0, viewport.height - GROUP_POPUP_PAD - desiredY)}px`;
    const position = clampPanelToViewport(tools, rect.left, desiredY, { pad: GROUP_POPUP_PAD });
    tools.style.left = `${Math.round(position.x)}px`;
    tools.style.top = `${Math.round(position.y)}px`;
  }
  const doc = group.ownerDocument;
  // Bubble phase: the chosen command runs first, then the popup goes.
  const choose = (event: Event): void => {
    if (event.target instanceof Element && choseLeaf(event.target)) closeGroupPopup(group);
  };
  popupListeners.get(group)?.();
  const detachDismiss = onOutsidePress(
    doc,
    () => [group],
    (event) => {
      if (event.target instanceof Element && event.target.closest(OPENED_LIST)) return;
      closeGroupPopup(group);
    },
    { event: 'pointerdown', capture: true },
  );
  doc.addEventListener('click', choose);
  popupListeners.set(group, () => {
    detachDismiss();
    doc.removeEventListener('click', choose);
  });
};

/** Fold `group` into one button named after it; the button opens the commands as a popup. */
const collapseGroup = (group: HTMLElement): void => {
  let button = group.querySelector<HTMLElement>(`:scope > .${COLLAPSE_BUTTON_CLASS}`);
  if (!button) {
    const name = group.getAttribute('aria-label') ?? '';
    const created = createRibbonButton({
      className: `fc-tb__rb fc-tb__rb--wide ${COLLAPSE_BUTTON_CLASS}`,
      text: name,
      title: name,
      ariaHaspopup: 'menu',
      ariaExpanded: false,
    });
    created.addEventListener('click', () => {
      if (group.classList.contains(GROUP_OPEN_CLASS)) closeGroupPopup(group);
      else openGroupPopup(group, created);
    });
    group.prepend(created);
    button = created;
  }
  group.classList.add(RIBBON_GROUP_COLLAPSED_CLASS);
};

/** Shrink the visible panel's groups, stage by stage and rightmost first, until it fits.
 *  Narrowing sets each tile label to its narrowest two-line width; condensing
 *  drops those widths for the row layout. */
export const fitRibbonPanel = (panel: HTMLElement, options: { collapse?: boolean } = {}): void => {
  const all = [...panel.querySelectorAll<HTMLElement>(':scope > .fc-tb__ribbon-group')];
  clearFit(all);
  if (panel.clientWidth === 0) return;
  for (const [stage, applies] of FIT_STAGES) {
    const groups = all.filter(applies);
    for (let i = groups.length - 1; i >= 0 && overflows(panel); i--) {
      const group = groups[i] as HTMLElement;
      const before = group.getBoundingClientRect().width;
      const labels = stage === RIBBON_GROUP_ICON_ONLY_CLASS ? [] : tileLabels(group);
      const saved = labels.map((label) => [label.style.width, label.style.maxWidth] as const);
      group.classList.add(stage);
      for (const label of labels) {
        if (stage === RIBBON_GROUP_NARROW_CLASS) fitLabelToTwoLines(label);
        else clearLabelWidth(label);
      }
      if (group.getBoundingClientRect().width >= before) {
        group.classList.remove(stage);
        labels.forEach((label, index) => {
          [label.style.width, label.style.maxWidth] = saved[index] ?? ['', ''];
        });
      }
    }
  }
  if (!options.collapse) return;
  const collapsible = all.filter(canCollapse);
  for (let i = collapsible.length - 1; i >= 0 && overflows(panel); i--) {
    collapseGroup(collapsible[i] as HTMLElement);
  }
};

/** Narrowest width in `[from, to]` at which `linesAt` reports at most two lines;
 *  `to` when none does. */
export const twoLineLabelWidth = (
  from: number,
  to: number,
  linesAt: (width: number) => number,
): number => {
  for (let width = Math.ceil(from); width < to; width++) {
    if (linesAt(width) <= 2) return width;
  }
  return Math.ceil(to);
};

const lineCount = (label: HTMLElement): number => {
  const range = label.ownerDocument.createRange();
  range.selectNodeContents(label);
  return new Set([...range.getClientRects()].map((rect) => Math.round(rect.top))).size;
};

const tileLabels = (group: HTMLElement): HTMLElement[] => [
  ...group.querySelectorAll<HTMLElement>(TILE_LABELS),
];

const clearLabelWidth = (label: HTMLElement): void => {
  label.style.removeProperty('width');
  label.style.removeProperty('max-width');
};

/** The fitted width overrides the stylesheet's label cap, so a label that
 *  needs more room than the cap still gets its two lines. */
const fitLabelToTwoLines = (label: HTMLElement): void => {
  label.style.maxWidth = 'none';
  label.style.width = 'min-content';
  const from = label.getBoundingClientRect().width;
  label.style.width = 'max-content';
  const to = label.getBoundingClientRect().width;
  if (from === 0 || to <= from) {
    clearLabelWidth(label);
    return;
  }
  const width = twoLineLabelWidth(from, to, (w) => {
    label.style.width = `${w}px`;
    return lineCount(label);
  });
  label.style.width = `${width}px`;
};

/** The nearest ancestor of `el` that lays out a box, skipping `display: contents` wrappers. */
const boxedAncestor = (el: HTMLElement): HTMLElement | null => {
  const view = el.ownerDocument.defaultView;
  let node = el.parentElement;
  while (node && view?.getComputedStyle(node).display === 'contents') node = node.parentElement;
  return node;
};

/** Refit the active panel under `root` now and whenever `root` changes width. */
export const attachRibbonFit = (root: HTMLElement): { refit(): void; detach(): void } => {
  const refit = (): void => {
    const shell = root.querySelector<HTMLElement>(':scope > .fc-tb__ribbon-shell');
    if (!shell) return;
    const singleLine = shell.classList.contains('fc-tb__ribbon-shell--singleLine');
    const groups = [...shell.querySelectorAll<HTMLElement>('.fc-tb__ribbon-group')];
    clearFit(groups);
    for (const label of groups.flatMap(tileLabels)) clearLabelWidth(label);
    if (root.dataset.fcPlatform === 'mac' && !singleLine) {
      for (const label of shell.querySelectorAll<HTMLElement>(TWO_LINE_LABELS)) {
        fitLabelToTwoLines(label);
      }
    }
    if (singleLine) return;
    const panel = shell.querySelector<HTMLElement>(':scope > .fc-tb__ribbon:not([hidden])');
    if (!panel) return;
    // A tile label too long for two lines at its default width widens until it fits.
    for (const group of panel.querySelectorAll<HTMLElement>(':scope > .fc-tb__ribbon-group')) {
      if (!isTileGroup(group)) continue;
      for (const label of tileLabels(group)) {
        if (lineCount(label) > 2) fitLabelToTwoLines(label);
      }
    }
    fitRibbonPanel(panel, { collapse: root.dataset.fcPlatform === 'mac' });
  };
  let lastWidth = -1;
  const observer =
    typeof ResizeObserver === 'undefined'
      ? null
      : new ResizeObserver((entries) => {
          const width = entries[0]?.contentRect.width ?? 0;
          if (width === lastWidth) return;
          lastWidth = width;
          refit();
        });
  observer?.observe(root);
  // A `display: contents` root (or wrapper) has no box and never reports a size change, so
  // the nearest ancestor that lays out a box stands in for it.
  const boxed = boxedAncestor(root);
  if (boxed) observer?.observe(boxed);
  // Label widths depend on the loaded web font, which can settle after mount.
  void root.ownerDocument.fonts?.ready.then(refit);
  return {
    refit,
    detach: () => observer?.disconnect(),
  };
};
