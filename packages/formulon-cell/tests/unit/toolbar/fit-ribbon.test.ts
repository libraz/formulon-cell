import { describe, expect, it } from 'vitest';
import {
  attachRibbonFit,
  fitRibbonPanel,
  RIBBON_GROUP_CONDENSED_CLASS,
  RIBBON_GROUP_ICON_ONLY_CLASS,
  RIBBON_GROUP_NARROW_CLASS,
  twoLineLabelWidth,
} from '../../../src/toolbar/ribbon/fit-ribbon.js';

type GroupSpec = {
  variant: string;
  /** Width at each stage: base, narrow, condensed, icon-only. */
  widths: readonly [number, number, number, number];
  buttons?: readonly string[];
};

const widthOf = (group: HTMLElement, spec: GroupSpec): number => {
  if (group.classList.contains(RIBBON_GROUP_ICON_ONLY_CLASS)) return spec.widths[3];
  if (group.classList.contains(RIBBON_GROUP_CONDENSED_CLASS)) return spec.widths[2];
  if (group.classList.contains(RIBBON_GROUP_NARROW_CLASS)) return spec.widths[1];
  return spec.widths[0];
};

/** A panel whose geometry follows the fit classes, since happy-dom does no layout. */
const makePanel = (clientWidth: number, specs: readonly GroupSpec[]) => {
  const panel = document.createElement('div');
  panel.className = 'fc-tb__ribbon';
  const groups = specs.map((spec) => {
    const group = document.createElement('section');
    group.className = `fc-tb__ribbon-group fc-tb__ribbon-group--${spec.variant}`;
    const tools = document.createElement('div');
    tools.className = 'fc-tb__ribbon-tools';
    for (const kind of spec.buttons ?? []) {
      const button = document.createElement('button');
      button.className = `fc-tb__rb fc-tb__rb--${kind}`;
      tools.appendChild(button);
    }
    group.appendChild(tools);
    group.getBoundingClientRect = () => ({ width: widthOf(group, spec) }) as DOMRect;
    panel.appendChild(group);
    return group;
  });
  Object.defineProperty(panel, 'clientWidth', { get: () => clientWidth });
  Object.defineProperty(panel, 'scrollWidth', {
    get: () => groups.reduce((sum, group, i) => sum + widthOf(group, specs[i] as GroupSpec), 0),
  });
  return { panel, groups };
};

const classesOf = (groups: readonly HTMLElement[]) =>
  groups.map((group) =>
    [RIBBON_GROUP_NARROW_CLASS, RIBBON_GROUP_CONDENSED_CLASS, RIBBON_GROUP_ICON_ONLY_CLASS]
      .filter((name) => group.classList.contains(name))
      .map((name) => name.replace('fc-tb__ribbon-group--', '')),
  );

describe('toolbar/ribbon fit', () => {
  it('leaves a panel that already fits untouched', () => {
    const { panel, groups } = makePanel(500, [
      { variant: 'tiles', widths: [200, 180, 150, 100] },
      { variant: 'tiles', widths: [200, 180, 150, 100] },
    ]);
    fitRibbonPanel(panel);
    expect(classesOf(groups)).toEqual([[], []]);
  });

  it('narrows tile groups from the right edge inward and stops once the panel fits', () => {
    const { panel, groups } = makePanel(370, [
      { variant: 'tiles', widths: [200, 180, 150, 100] },
      { variant: 'tiles', widths: [200, 180, 150, 100] },
    ]);
    fitRibbonPanel(panel);
    expect(classesOf(groups)).toEqual([['narrow'], ['narrow']]);

    const fitsAfterOne = makePanel(390, [
      { variant: 'tiles', widths: [200, 180, 150, 100] },
      { variant: 'tiles', widths: [200, 180, 150, 100] },
    ]);
    fitRibbonPanel(fitsAfterOne.panel);
    expect(classesOf(fitsAfterOne.groups)).toEqual([[], ['narrow']]);
  });

  it('condenses after narrowing and keeps a stage only where it saved width', () => {
    const { panel, groups } = makePanel(250, [
      { variant: 'tiles', widths: [200, 180, 150, 150] },
      // A single tile whose label is longer on one line than wrapped.
      { variant: 'tiles', widths: [120, 100, 130, 130] },
    ]);
    fitRibbonPanel(panel);
    expect(classesOf(groups)).toEqual([['narrow', 'condensed'], ['narrow']]);
  });

  it('drops row labels last, skipping groups with text-only buttons', () => {
    const { panel, groups } = makePanel(200, [
      { variant: 'clipboard', widths: [80, 80, 80, 80], buttons: ['large'] },
      { variant: 'compact', widths: [150, 150, 150, 60], buttons: ['stacked'] },
      { variant: 'tiles', widths: [150, 150, 120, 60], buttons: ['wide', 'mono'] },
    ]);
    fitRibbonPanel(panel);
    expect(classesOf(groups)).toEqual([[], ['icon-only'], ['condensed']]);
  });

  it('clears every stage when the shell switches to the single-line ribbon', () => {
    const root = document.createElement('div');
    const shell = document.createElement('div');
    shell.className = 'fc-tb__ribbon-shell fc-tb__ribbon-shell--singleLine';
    const { panel, groups } = makePanel(100, [{ variant: 'tiles', widths: [200, 180, 150, 100] }]);
    groups[0]?.classList.add(RIBBON_GROUP_NARROW_CLASS, RIBBON_GROUP_CONDENSED_CLASS);
    shell.appendChild(panel);
    root.appendChild(shell);

    const fit = attachRibbonFit(root);
    fit.refit();
    fit.detach();
    expect(classesOf(groups)).toEqual([[]]);
  });
});

describe('toolbar/ribbon two-line label width', () => {
  // A label of 120px set in words; lines = ceil(120 / width) stands in for layout.
  const linesAt = (width: number): number => Math.ceil(120 / width);

  it('returns the narrowest width that keeps the label on two lines', () => {
    expect(twoLineLabelWidth(54, 120, linesAt)).toBe(60);
  });

  it('keeps the starting width when the label already fits on two lines', () => {
    expect(twoLineLabelWidth(64, 120, linesAt)).toBe(64);
  });

  it('falls back to the single-line width when no narrower width fits', () => {
    expect(twoLineLabelWidth(54, 90, () => 3)).toBe(90);
  });
});
