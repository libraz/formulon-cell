import { describe, expect, it } from 'vitest';
import {
  attachRibbonFit,
  fitRibbonPanel,
  RIBBON_GROUP_COLLAPSED_CLASS,
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

describe('toolbar/ribbon refit on resize', () => {
  it('watches the root and its parent, since a display: contents root never reports a size', () => {
    const observed: Element[] = [];
    const original = globalThis.ResizeObserver;
    globalThis.ResizeObserver = class {
      observe(el: Element): void {
        observed.push(el);
      }
      disconnect(): void {}
      unobserve(): void {}
    } as unknown as typeof ResizeObserver;
    try {
      const parent = document.createElement('div');
      const root = document.createElement('div');
      parent.appendChild(root);
      const fit = attachRibbonFit(root);
      expect(observed).toEqual([root, parent]);
      fit.detach();
    } finally {
      globalThis.ResizeObserver = original;
    }
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

describe('toolbar/ribbon group collapse', () => {
  const build = (initialWidth: number) => {
    let clientWidth = initialWidth;
    const panel = document.createElement('div');
    panel.className = 'fc-tb__ribbon';
    const groups = ['クリップボード', 'フォント', 'スタイル', '編集'].map((name, index) => {
      const group = document.createElement('section');
      group.className = `fc-tb__ribbon-group${index === 2 ? ' fc-tb__ribbon-group--styles' : ''}`;
      group.setAttribute('aria-label', name);
      const tools = document.createElement('div');
      tools.className = 'fc-tb__ribbon-tools';
      for (let i = 0; i < 3; i += 1) {
        const button = document.createElement('button');
        button.className = 'fc-tb__rb';
        button.dataset.ribbonCommand = `cmd-${index}-${i}`;
        tools.appendChild(button);
      }
      group.appendChild(tools);
      group.getBoundingClientRect = () =>
        ({
          width: group.classList.contains(RIBBON_GROUP_COLLAPSED_CLASS) ? 48 : 200,
          left: 0,
          bottom: 0,
        }) as DOMRect;
      panel.appendChild(group);
      return group;
    });
    Object.defineProperty(panel, 'clientWidth', { get: () => clientWidth });
    Object.defineProperty(panel, 'scrollWidth', {
      get: () => groups.reduce((sum, group) => sum + group.getBoundingClientRect().width, 0),
    });
    document.body.appendChild(panel);
    return {
      panel,
      groups,
      setWidth: (width: number) => {
        clientWidth = width;
      },
    };
  };

  it('does nothing unless asked to collapse', () => {
    const { panel, groups } = build(300);
    fitRibbonPanel(panel);
    expect(groups.some((group) => group.classList.contains(RIBBON_GROUP_COLLAPSED_CLASS))).toBe(
      false,
    );
  });

  it('folds groups into named buttons from the right, sparing the style gallery group', () => {
    const { panel, groups } = build(520);
    fitRibbonPanel(panel, { collapse: true });
    expect(groups.map((group) => group.classList.contains(RIBBON_GROUP_COLLAPSED_CLASS))).toEqual([
      false,
      true,
      false,
      true,
    ]);
    const button = groups[3]?.querySelector<HTMLButtonElement>('.fc-tb__ribbon-group-collapse');
    expect(button?.textContent).toBe('編集');
    button?.click();
    expect(groups[3]?.classList.contains('fc-tb__ribbon-group--open')).toBe(true);
    expect(button?.getAttribute('aria-expanded')).toBe('true');
    groups[3]?.querySelector<HTMLElement>('[data-ribbon-command]')?.click();
    expect(groups[3]?.classList.contains('fc-tb__ribbon-group--open')).toBe(false);
  });

  it('keeps a folded group popup open for dropdowns and their lists, and inside the viewport', () => {
    const { panel, groups } = build(520);
    fitRibbonPanel(panel, { collapse: true });
    const group = groups[3] as HTMLElement;
    const tools = group.querySelector<HTMLElement>('.fc-tb__ribbon-tools') as HTMLElement;
    const dropdown = document.createElement('button');
    dropdown.dataset.ribbonCommand = 'fontFamily';
    dropdown.setAttribute('aria-haspopup', 'listbox');
    const option = document.createElement('div');
    option.className = 'fc-tb__rb-dd__list';
    tools.append(dropdown, option);

    const previousWidth = window.innerWidth;
    window.innerWidth = 390;
    window.innerHeight = 800;
    tools.getBoundingClientRect = () => ({ width: 600, height: 100 }) as DOMRect;
    const open = group.querySelector<HTMLButtonElement>('.fc-tb__ribbon-group-collapse');
    open?.click();
    window.innerWidth = previousWidth;

    expect(tools.style.maxWidth).toBe('374px');
    expect(tools.style.left).toBe('8px');
    expect(tools.style.maxHeight).toBe('790px');

    dropdown.click();
    option.click();
    option.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    expect(group.classList.contains('fc-tb__ribbon-group--open')).toBe(true);

    document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    expect(group.classList.contains('fc-tb__ribbon-group--open')).toBe(false);
    expect(tools.style.maxWidth).toBe('');
  });

  it('expands every group again once the panel fits', () => {
    const { panel, groups, setWidth } = build(450);
    fitRibbonPanel(panel, { collapse: true });
    setWidth(2000);
    fitRibbonPanel(panel, { collapse: true });
    expect(groups.some((group) => group.classList.contains(RIBBON_GROUP_COLLAPSED_CLASS))).toBe(
      false,
    );
  });
});
