/**
 * Width fitting for the ribbon panel.
 *
 * When the visible panel is wider than its viewport, tile groups give up width
 * in stages, each applied from the right edge inward: tiles narrow so labels
 * wrap onto a second line, then large buttons drop to icon-and-label rows
 * stacked three high, and finally those rows drop their labels (the accessible
 * name and tooltip stay). Command order never changes; only density does. A
 * stage is kept on a group only if it made the group narrower.
 */

export const RIBBON_GROUP_NARROW_CLASS = 'fc-tb__ribbon-group--narrow';
export const RIBBON_GROUP_CONDENSED_CLASS = 'fc-tb__ribbon-group--condensed';
export const RIBBON_GROUP_ICON_ONLY_CLASS = 'fc-tb__ribbon-group--icon-only';

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

const overflows = (panel: HTMLElement): boolean => panel.scrollWidth > panel.clientWidth + 1;

const clearFit = (groups: Iterable<Element>): void => {
  for (const group of groups) group.classList.remove(...FIT_CLASSES);
};

/** Shrink the visible panel's groups, stage by stage and rightmost first, until it fits. */
export const fitRibbonPanel = (panel: HTMLElement): void => {
  const all = [...panel.querySelectorAll<HTMLElement>(':scope > .fc-tb__ribbon-group')];
  clearFit(all);
  if (panel.clientWidth === 0) return;
  for (const [stage, applies] of FIT_STAGES) {
    const groups = all.filter(applies);
    for (let i = groups.length - 1; i >= 0 && overflows(panel); i--) {
      const group = groups[i] as HTMLElement;
      const before = group.getBoundingClientRect().width;
      group.classList.add(stage);
      if (group.getBoundingClientRect().width >= before) group.classList.remove(stage);
    }
  }
};

/** Refit the active panel under `root` now and whenever `root` changes width. */
export const attachRibbonFit = (root: HTMLElement): { refit(): void; detach(): void } => {
  const refit = (): void => {
    const shell = root.querySelector<HTMLElement>(':scope > .fc-tb__ribbon-shell');
    if (!shell) return;
    if (shell.classList.contains('fc-tb__ribbon-shell--singleLine')) {
      clearFit(shell.querySelectorAll('.fc-tb__ribbon-group'));
      return;
    }
    const panel = shell.querySelector<HTMLElement>(':scope > .fc-tb__ribbon:not([hidden])');
    if (panel) fitRibbonPanel(panel);
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
  return {
    refit,
    detach: () => observer?.disconnect(),
  };
};
