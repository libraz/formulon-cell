import { describe, expect, it, vi } from 'vitest';
import { pageScaleMenuText } from '../../../src/toolbar/menu-text.js';
import { createControlDispatch } from '../../../src/toolbar/ribbon/control-dispatch.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';

const createIcon = () =>
  createControlDispatch({
    getInst: () => null,
    ribbonLang: 'ja',
    ribbonText: toolbarText('ja'),
    pageScaleText: pageScaleMenuText('ja'),
    sheetEl: document.createElement('div'),
    focusSheet: vi.fn(),
    refreshWorkbookCells: vi.fn(),
    projectFormatToolbar: vi.fn(),
  }).createRibbonIcon;

describe('toolbar/ribbon control dispatch icons', () => {
  it('renders Excel-like multi-color SVG segments before Fluent paths', () => {
    const svg = createIcon()('fillColor');

    expect(svg?.classList.contains('demo__rb-icon')).toBe(true);
    expect(svg?.getAttribute('viewBox')).toBe('0 0 24 24');
    expect(svg?.getAttribute('fill')).toBeNull();

    const paths = Array.from(svg?.querySelectorAll('path') ?? []);
    expect(paths.length).toBeGreaterThan(1);
    expect(paths.map((path) => path.getAttribute('fill'))).toContain('#ffd966');
    expect(paths.some((path) => path.hasAttribute('stroke'))).toBe(true);
  });

  it('renders Excel-like Clipboard SVGs for high-visibility Home ribbon commands', () => {
    for (const icon of ['paste', 'cut', 'copy', 'paint']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
    }
  });

  it('keeps the Cut ribbon icon readable as scissors', () => {
    const svg = createIcon()('cut');
    const paths = Array.from(svg?.querySelectorAll('path') ?? []);
    const handlePaths = paths.filter((path) => path.getAttribute('fill') === '#2f75b5');
    const handleHolePaths = paths.filter((path) => path.getAttribute('fill') === '#ffffff');
    const bladePaths = paths.filter((path) => path.getAttribute('stroke') === '#1f1f1f');

    expect(handlePaths).toHaveLength(2);
    expect(handleHolePaths).toHaveLength(2);
    expect(bladePaths.length).toBeGreaterThanOrEqual(2);
    expect(svg?.querySelector('path[fill="#107c41"]')).toBeTruthy();
    expect(handlePaths.every((path) => path.getAttribute('d')?.includes('a2.6'))).toBe(true);
    expect(handleHolePaths.every((path) => path.getAttribute('d')?.includes('a1.25'))).toBe(true);
    expect(paths.some((path) => path.getAttribute('stroke') === '#1f1f1f')).toBe(true);
  });

  it('keeps Copy and Format Painter icons readable as their source figures', () => {
    const copyPaths = Array.from(createIcon()('copy')?.querySelectorAll('path') ?? []);
    expect(copyPaths.filter((path) => path.getAttribute('fill') === '#f3f8ff')).toHaveLength(1);
    expect(copyPaths.filter((path) => path.getAttribute('fill') === '#ffffff')).toHaveLength(1);
    expect(copyPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
    expect(copyPaths.some((path) => path.getAttribute('stroke') === '#8a8f98')).toBe(true);

    const paintPaths = Array.from(createIcon()('paint')?.querySelectorAll('path') ?? []);
    expect(paintPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(paintPaths.some((path) => path.getAttribute('fill') === '#f4c27a')).toBe(true);
    expect(paintPaths.some((path) => path.getAttribute('fill') === '#f3f2f1')).toBe(true);
    expect(paintPaths.some((path) => path.getAttribute('fill') === '#2f75b5')).toBe(true);
    expect(paintPaths.some((path) => path.getAttribute('stroke') === '#ffffff')).toBe(true);
  });

  it('renders Excel-like Font size SVGs with semantic A and arrow shapes', () => {
    for (const icon of ['fontGrow', 'fontShrink']) {
      const svg = createIcon()(icon);
      const arrowPath = Array.from(svg?.querySelectorAll('path') ?? []).find(
        (path) => path.getAttribute('stroke') === '#107c41',
      );

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(
        svg?.querySelector('path[stroke], path[fill="#107c41"], path[fill="#1f1f1f"]'),
      ).toBeTruthy();
      expect(arrowPath?.getAttribute('stroke-width')).toBe('1.75');
      expect(arrowPath?.getAttribute('stroke-linejoin')).toBe('round');
    }
  });

  it('renders Excel-like text emphasis SVGs for compact Font group toggles', () => {
    for (const icon of ['bold', 'italic', 'underline', 'strike']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelector('path[fill="#1f1f1f"]')).toBeTruthy();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
    }

    expect(createIcon()('underline')?.querySelector('path[fill="#107c41"]')).toBeTruthy();
    expect(createIcon()('strike')?.querySelector('path[fill="#107c41"]')).toBeTruthy();
  });

  it('renders Excel-like Number group SVGs with semantic numeric format shapes', () => {
    for (const icon of ['currency', 'percent', 'comma', 'decDown', 'decUp']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(
        svg?.querySelector('path[stroke], path[fill="#107c41"], path[fill="#1f1f1f"]'),
      ).toBeTruthy();
    }
  });

  it('keeps compact Number group icons readable as symbols, separators, and decimal arrows', () => {
    const currencyPaths = Array.from(createIcon()('currency')?.querySelectorAll('path') ?? []);
    expect(currencyPaths.some((path) => path.getAttribute('fill') === '#eef6ee')).toBe(true);
    expect(currencyPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(currencyPaths.some((path) => path.getAttribute('d')?.includes('M9.4 9.4 12 12.2'))).toBe(
      true,
    );

    const percentPaths = Array.from(createIcon()('percent')?.querySelectorAll('path') ?? []);
    expect(percentPaths.filter((path) => path.getAttribute('stroke') === '#1f1f1f')).toHaveLength(
      3,
    );
    expect(percentPaths.every((path) => path.getAttribute('stroke-width'))).toBe(true);

    const commaPaths = Array.from(createIcon()('comma')?.querySelectorAll('path') ?? []);
    expect(commaPaths.some((path) => path.getAttribute('fill') === '#1f1f1f')).toBe(true);
    expect(
      commaPaths.some((path) => path.getAttribute('d')?.includes('c.5 2-.4 3.6-2.5 4.6')),
    ).toBe(true);
    expect(commaPaths.some((path) => path.getAttribute('fill') === '#2f75b5')).toBe(false);

    for (const icon of ['decDown', 'decUp']) {
      const paths = Array.from(createIcon()(icon)?.querySelectorAll('path') ?? []);
      const arrowPath = paths.find((path) => path.getAttribute('stroke') === '#107c41');

      expect(paths.some((path) => path.getAttribute('stroke') === '#1f1f1f')).toBe(true);
      expect(arrowPath).toBeTruthy();
      expect(arrowPath?.getAttribute('stroke-width')).toBe('1.8');
    }
  });

  it('renders readable Excel-like Alignment group SVGs with semantic guide and text marks', () => {
    for (const icon of [
      'top',
      'middle',
      'bottomAlign',
      'alignLeft',
      'alignCenter',
      'alignRight',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]) {
      const svg = createIcon()(icon);
      const paths = Array.from(svg?.querySelectorAll('path') ?? []);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(paths.length).toBeGreaterThan(1);
      expect(
        paths.some(
          (path) =>
            path.getAttribute('stroke') === '#107c41' || path.getAttribute('fill') === '#107c41',
        ),
      ).toBe(true);
      expect(
        paths.some(
          (path) =>
            path.getAttribute('fill') === '#1f1f1f' || path.getAttribute('stroke') === '#1f1f1f',
        ),
      ).toBe(true);
    }
  });

  it('keeps compact Alignment glyphs bold enough to remain legible at ribbon size', () => {
    for (const icon of ['alignLeft', 'alignCenter', 'alignRight']) {
      const paths = Array.from(createIcon()(icon)?.querySelectorAll('path') ?? []);
      const guide = paths.find((path) => path.getAttribute('fill') === '#107c41');
      const textRows = paths.find((path) => path.getAttribute('fill') === '#1f1f1f');

      expect(paths).toHaveLength(2);
      expect(guide?.getAttribute('d')).toMatch(/h3v15\.4/);
      expect(textRows?.getAttribute('d')).toContain('v2.8');
      expect(textRows?.getAttribute('stroke')).toBeNull();
    }

    for (const icon of ['indentDecrease', 'indentIncrease']) {
      const paths = Array.from(createIcon()(icon)?.querySelectorAll('path') ?? []);
      const arrow = paths.find((path) => path.getAttribute('stroke') === '#107c41');
      const textRows = paths.find((path) => path.getAttribute('fill') === '#1f1f1f');

      expect(textRows?.getAttribute('d')).toContain('v2.8');
      expect(arrow?.getAttribute('stroke-width')).toBe('2.65');
      expect(arrow?.getAttribute('d')).toMatch(/h5\.[68]/);
    }

    const wrapPaths = Array.from(createIcon()('wrap')?.querySelectorAll('path') ?? []);
    expect(
      wrapPaths.find((path) => path.getAttribute('fill') === '#1f1f1f')?.getAttribute('d'),
    ).toContain('V19');
    expect(
      wrapPaths
        .find((path) => path.getAttribute('stroke') === '#107c41')
        ?.getAttribute('stroke-width'),
    ).toBe('2.6');

    const orientationPaths = Array.from(
      createIcon()('textOrientation')?.querySelectorAll('path') ?? [],
    );
    expect(
      orientationPaths.find((path) => path.getAttribute('fill') === '#1f1f1f')?.getAttribute('d'),
    ).toContain('4.8h3');
    expect(
      orientationPaths
        .find((path) => path.getAttribute('stroke') === '#107c41')
        ?.getAttribute('stroke-width'),
    ).toBe('2.4');

    const mergePaths = Array.from(createIcon()('merge')?.querySelectorAll('path') ?? []);
    expect(mergePaths.find((path) => path.getAttribute('fill') === '#d9d9d9')).toBeTruthy();
    expect(
      mergePaths.find((path) => path.getAttribute('stroke') === '#107c41')?.getAttribute('d'),
    ).toContain('M7.5 12h9');
    expect(
      mergePaths
        .find((path) => path.getAttribute('stroke') === '#107c41')
        ?.getAttribute('stroke-width'),
    ).toBe('2.35');
  });

  it('keeps Home Styles group icons visually distinct as rules, table, and style swatches', () => {
    const conditionalPaths = Array.from(
      createIcon()('conditional')?.querySelectorAll('path') ?? [],
    );
    expect(conditionalPaths.some((path) => path.getAttribute('fill') === '#c00000')).toBe(true);
    expect(conditionalPaths.some((path) => path.getAttribute('fill') === '#ffd966')).toBe(true);
    expect(conditionalPaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);

    const tablePaths = Array.from(createIcon()('tableStyle')?.querySelectorAll('path') ?? []);
    expect(tablePaths.some((path) => path.getAttribute('fill') === '#2f75b5')).toBe(true);
    expect(tablePaths.some((path) => path.getAttribute('fill') === '#ed7d31')).toBe(true);
    expect(tablePaths.some((path) => path.getAttribute('stroke') === '#8f4a12')).toBe(true);

    const cellStylePaths = Array.from(createIcon()('cellStyles')?.querySelectorAll('path') ?? []);
    for (const fill of ['#e2f0d9', '#fff2cc', '#ddebf7', '#fce4d6']) {
      expect(cellStylePaths.some((path) => path.getAttribute('fill') === fill)).toBe(true);
    }
    expect(cellStylePaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(cellStylePaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
  });

  it('keeps Home Cells group icons visually distinct as insert, delete, and format actions', () => {
    const insertPaths = Array.from(createIcon()('insertRows')?.querySelectorAll('path') ?? []);
    expect(insertPaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);
    expect(insertPaths.some((path) => path.getAttribute('stroke') === '#ffffff')).toBe(true);
    expect(insertPaths.some((path) => path.getAttribute('d')?.includes('M6 4.5v3'))).toBe(true);

    const deletePaths = Array.from(createIcon()('deleteRows')?.querySelectorAll('path') ?? []);
    expect(deletePaths.some((path) => path.getAttribute('fill') === '#c00000')).toBe(true);
    expect(deletePaths.some((path) => path.getAttribute('stroke') === '#ffffff')).toBe(true);
    expect(deletePaths.some((path) => path.getAttribute('d')?.includes('M4.9 4.9'))).toBe(true);

    const formatPaths = Array.from(createIcon()('formatCells')?.querySelectorAll('path') ?? []);
    expect(formatPaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);
    expect(formatPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
    expect(formatPaths.some((path) => path.getAttribute('d')?.includes('M15.8 14.4v1.2'))).toBe(
      true,
    );
  });

  it('renders Excel-like Insert table SVGs with distinct table and PivotTable shapes', () => {
    for (const icon of ['table', 'pivotTable']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    const tablePaths = Array.from(createIcon()('table')?.querySelectorAll('path') ?? []);
    expect(tablePaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);
    expect(tablePaths.some((path) => path.getAttribute('stroke') === '#ffffff')).toBe(true);

    const pivotPaths = Array.from(createIcon()('pivotTable')?.querySelectorAll('path') ?? []);
    expect(pivotPaths.some((path) => path.getAttribute('fill') === '#2f75b5')).toBe(true);
    expect(pivotPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(pivotPaths.some((path) => path.getAttribute('fill') === '#eef6ee')).toBe(true);
    expect(pivotPaths.some((path) => path.getAttribute('d')?.includes('19.2 11.2l1.8'))).toBe(true);
  });

  it('renders Excel-like Insert illustration SVGs with semantic media shapes', () => {
    for (const icon of ['picture', 'shapes', 'screenshot']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    const picturePaths = Array.from(createIcon()('picture')?.querySelectorAll('path') ?? []);
    expect(picturePaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);
    expect(picturePaths.some((path) => path.getAttribute('fill') === '#ffd966')).toBe(true);
    expect(picturePaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);

    const shapePaths = Array.from(createIcon()('shapes')?.querySelectorAll('path') ?? []);
    expect(shapePaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
    expect(shapePaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(shapePaths.some((path) => path.getAttribute('fill') === '#ffd966')).toBe(true);
    expect(shapePaths.some((path) => path.getAttribute('fill') === '#ed7d31')).toBe(true);

    const screenshotPaths = Array.from(createIcon()('screenshot')?.querySelectorAll('path') ?? []);
    expect(screenshotPaths.some((path) => path.getAttribute('fill') === '#eaf2fb')).toBe(true);
    expect(screenshotPaths.some((path) => path.getAttribute('fill') === '#d7e9fb')).toBe(true);
    expect(
      screenshotPaths.some(
        (path) =>
          path.getAttribute('stroke') === '#107c41' && path.getAttribute('stroke-width') === '1.65',
      ),
    ).toBe(true);
  });

  it('renders Excel-like Insert command SVGs with semantic chart/link/comment/symbol shapes', () => {
    for (const icon of ['chart', 'link', 'commentAdd', 'function']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }
  });

  it('renders Excel-like Page Layout SVGs with semantic page and print shapes', () => {
    for (const icon of [
      'pageTheme',
      'pageSetup',
      'printArea',
      'pageBreaks',
      'sheetBackground',
      'printTitles',
    ]) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    const pageSetupPaths = Array.from(createIcon()('pageSetup')?.querySelectorAll('path') ?? []);
    expect(pageSetupPaths.some((path) => path.getAttribute('fill') === '#eef6ee')).toBe(true);
    expect(
      pageSetupPaths.some(
        (path) =>
          path.getAttribute('stroke') === '#107c41' &&
          path.getAttribute('d')?.includes('17.2 13.6'),
      ),
    ).toBe(true);
    expect(pageSetupPaths.some((path) => path.getAttribute('stroke') === '#0b5a2f')).toBe(true);

    const pageBreakPaths = Array.from(createIcon()('pageBreaks')?.querySelectorAll('path') ?? []);
    expect(
      pageBreakPaths.some(
        (path) =>
          path.getAttribute('stroke') === '#2f75b5' &&
          path.getAttribute('stroke-dasharray') === '2 1.4',
      ),
    ).toBe(true);
    expect(pageBreakPaths.filter((path) => path.getAttribute('stroke-dasharray')).length).toBe(2);
  });

  it('renders Excel-like Data tab SVGs with semantic data operation shapes', () => {
    for (const icon of [
      'filter',
      'textToColumns',
      'removeDuplicates',
      'dataValidation',
      'outlineGroup',
      'outlineUngroup',
      'outlineShow',
      'outlineHide',
    ]) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    const removeDuplicatePaths = Array.from(
      createIcon()('removeDuplicates')?.querySelectorAll('path') ?? [],
    );
    expect(
      removeDuplicatePaths.filter((path) => path.getAttribute('fill') === '#ffffff'),
    ).toHaveLength(2);
    expect(removeDuplicatePaths.some((path) => path.getAttribute('stroke') === '#c00000')).toBe(
      true,
    );

    const validationPaths = Array.from(
      createIcon()('dataValidation')?.querySelectorAll('path') ?? [],
    );
    expect(validationPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(validationPaths.some((path) => path.getAttribute('fill') === '#ffd966')).toBe(true);

    const outlineGroupPaths = Array.from(
      createIcon()('outlineGroup')?.querySelectorAll('path') ?? [],
    );
    const outlineUngroupPaths = Array.from(
      createIcon()('outlineUngroup')?.querySelectorAll('path') ?? [],
    );
    expect(outlineGroupPaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);
    expect(outlineUngroupPaths.some((path) => path.getAttribute('fill') === '#c00000')).toBe(true);
    expect(createIcon()('outlineShow')?.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(createIcon()('outlineHide')?.querySelector('path[stroke="#c00000"]')).toBeTruthy();
  });

  it('renders Excel-like Formulas and Review SVGs with semantic audit and proofing shapes', () => {
    for (const icon of [
      'names',
      'trace',
      'dependents',
      'clearArrows',
      'errorChecking',
      'calcOptions',
      'watch',
      'spelling',
      'accessibility',
      'translate',
      'protect',
    ]) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    const tracePaths = Array.from(createIcon()('trace')?.querySelectorAll('path') ?? []);
    const dependentPaths = Array.from(createIcon()('dependents')?.querySelectorAll('path') ?? []);
    expect(tracePaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(tracePaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
    expect(dependentPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(dependentPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);

    const clearArrowPaths = Array.from(createIcon()('clearArrows')?.querySelectorAll('path') ?? []);
    expect(clearArrowPaths.some((path) => path.getAttribute('stroke') === '#8a8f98')).toBe(true);
    expect(clearArrowPaths.some((path) => path.getAttribute('stroke') === '#c00000')).toBe(true);

    const errorPaths = Array.from(createIcon()('errorChecking')?.querySelectorAll('path') ?? []);
    expect(errorPaths.some((path) => path.getAttribute('fill') === '#ffd966')).toBe(true);
    expect(errorPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);

    const watchPaths = Array.from(createIcon()('watch')?.querySelectorAll('path') ?? []);
    expect(watchPaths.some((path) => path.getAttribute('fill') === '#eef6ee')).toBe(true);
    expect(watchPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);

    expect(createIcon()('spelling')?.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(createIcon()('accessibility')?.querySelector('path[fill="#2f75b5"]')).toBeTruthy();
    expect(createIcon()('translate')?.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(createIcon()('translate')?.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(createIcon()('protect')?.querySelector('path[fill="#ffd966"]')).toBeTruthy();
  });

  it('renders Excel-like File and View SVGs with semantic page, print, freeze, and zoom shapes', () => {
    for (const icon of ['page', 'print', 'freeze', 'zoom']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    expect(createIcon()('page')?.querySelector('path[fill="#eef6ee"]')).toBeTruthy();
    expect(createIcon()('print')?.querySelector('path[fill="#eef6ee"]')).toBeTruthy();

    const freezePaths = Array.from(createIcon()('freeze')?.querySelectorAll('path') ?? []);
    expect(freezePaths.some((path) => path.getAttribute('fill') === '#dceef8')).toBe(true);
    expect(freezePaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);

    const zoomPaths = Array.from(createIcon()('zoom')?.querySelectorAll('path') ?? []);
    expect(zoomPaths.some((path) => path.getAttribute('stroke') === '#1f1f1f')).toBe(true);
    expect(zoomPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
  });

  it('renders Excel-like optional and generic command SVGs with semantic tool shapes', () => {
    for (const icon of [
      'goTo',
      'goToSpecial',
      'options',
      'pen',
      'eraser',
      'script',
      'addIn',
      'pdf',
      'save',
      'saveAs',
      'autosave',
      'share',
      'pivotRecommended',
      'pivotExistingSheet',
      'dataValidationCircle',
      'dataValidationClearCircles',
      'dataValidationClearRules',
      'namesCreateTop',
      'namesCreateBottom',
      'namesCreateLeft',
      'namesCreateRight',
    ]) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    expect(createIcon()('goTo')?.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(createIcon()('goToSpecial')?.querySelector('path[fill="#d9c2f0"]')).toBeTruthy();

    const optionsPaths = Array.from(createIcon()('options')?.querySelectorAll('path') ?? []);
    expect(optionsPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
    expect(optionsPaths.some((path) => path.getAttribute('stroke-dasharray') === '1.9 3.05')).toBe(
      true,
    );

    expect(createIcon()('pdf')?.querySelector('path[fill="#c00000"]')).toBeTruthy();
    expect(createIcon()('save')?.querySelector('path[fill="#2f75b5"]')).toBeTruthy();
    expect(createIcon()('saveAs')?.querySelector('path[fill="#ffd966"]')).toBeTruthy();
    expect(createIcon()('autosave')?.querySelector('path[fill="#107c41"]')).toBeTruthy();
    expect(createIcon()('share')?.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(createIcon()('pivotRecommended')?.querySelector('path[fill="#ffd966"]')).toBeTruthy();
    expect(createIcon()('pivotExistingSheet')?.querySelector('path[fill="#107c41"]')).toBeTruthy();
  });

  it('renders Excel-like visual tile SVGs with chart, picture, shape, and screenshot semantics', () => {
    for (const icon of [
      'chartColumn',
      'chartBar',
      'chartLine',
      'chartArea',
      'chartPie',
      'chartScatter',
      'chartRecommended',
      'devicePicture',
      'onlinePicture',
      'stockPicture',
      'screenshotWindow',
      'screenClipping',
      'shapeLine',
      'shapeArrow',
      'shapeRectangle',
      'shapeRoundedRectangle',
      'shapeOval',
      'shapeTriangle',
      'shapeDiamond',
      'themeLight',
      'themeDark',
      'themeContrast',
    ]) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(0);
    }

    expect(createIcon()('chartRecommended')?.querySelector('path[fill="#ffd966"]')).toBeTruthy();
    expect(createIcon()('devicePicture')?.querySelector('path[fill="#107c41"]')).toBeTruthy();
    expect(createIcon()('onlinePicture')?.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(createIcon()('stockPicture')?.querySelector('path[fill="#f3f8ff"]')).toBeTruthy();

    expect(createIcon()('screenshotWindow')?.querySelector('path[fill="#f3f8ff"]')).toBeTruthy();
    expect(
      createIcon()('screenClipping')?.querySelector('path[stroke-dasharray="2 1"]'),
    ).toBeTruthy();
    expect(createIcon()('screenClipping')?.querySelector('path[fill="#107c41"]')).toBeTruthy();

    expect(createIcon()('shapeArrow')?.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(createIcon()('shapeRectangle')?.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
    expect(createIcon()('shapeTriangle')?.querySelector('path[stroke="#107c41"]')).toBeTruthy();
    expect(createIcon()('shapeDiamond')?.querySelector('path[stroke="#2f75b5"]')).toBeTruthy();
  });

  it('renders Excel-like Editing group SVGs with semantic multi-part shapes', () => {
    for (const icon of ['autosum', 'fill', 'clear', 'sortFilter', 'find']) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(svg?.querySelector('path[stroke]')).toBeTruthy();
    }

    const autosumPaths = Array.from(createIcon()('autosum')?.querySelectorAll('path') ?? []);
    expect(autosumPaths.some((path) => path.getAttribute('fill') === '#2f75b5')).toBe(true);
    expect(autosumPaths.some((path) => path.getAttribute('d')?.includes('5.9 4.7'))).toBe(true);

    const fillPaths = Array.from(createIcon()('fill')?.querySelectorAll('path') ?? []);
    expect(fillPaths.some((path) => path.getAttribute('stroke') === '#107c41')).toBe(true);
    expect(fillPaths.some((path) => path.getAttribute('d')?.includes('15.3 14.2 18 17'))).toBe(
      true,
    );

    const clearPaths = Array.from(createIcon()('clear')?.querySelectorAll('path') ?? []);
    expect(clearPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
    expect(clearPaths.some((path) => path.getAttribute('d')?.includes('13.2 7l4.8 4.8'))).toBe(
      true,
    );

    const sortFilterPaths = Array.from(createIcon()('sortFilter')?.querySelectorAll('path') ?? []);
    expect(sortFilterPaths.some((path) => path.getAttribute('fill') === '#2f75b5')).toBe(true);
    expect(sortFilterPaths.some((path) => path.getAttribute('fill') === '#107c41')).toBe(true);
    expect(sortFilterPaths.some((path) => path.getAttribute('stroke') === '#0b5a2f')).toBe(true);

    const findPaths = Array.from(createIcon()('find')?.querySelectorAll('path') ?? []);
    expect(findPaths.some((path) => path.getAttribute('stroke') === '#1f1f1f')).toBe(true);
    expect(findPaths.some((path) => path.getAttribute('stroke') === '#2f75b5')).toBe(true);
  });

  it('renders Excel-like Alignment group SVGs with semantic multi-part shapes', () => {
    for (const icon of [
      'top',
      'middle',
      'bottomAlign',
      'alignLeft',
      'alignCenter',
      'alignRight',
      'textOrientation',
      'wrap',
      'indentDecrease',
      'indentIncrease',
      'merge',
    ]) {
      const svg = createIcon()(icon);

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(svg?.querySelectorAll('path').length).toBeGreaterThan(1);
      expect(
        svg?.querySelector('path[stroke], path[fill="#107c41"], path[fill="#1f1f1f"]'),
      ).toBeTruthy();
    }
  });

  it('keeps compact Alignment icons readable with green guides and clear text shapes', () => {
    for (const icon of ['top', 'middle', 'bottomAlign']) {
      const svg = createIcon()(icon);
      const paths = Array.from(svg?.querySelectorAll('path') ?? []);
      const guidePath = paths.find((path) => path.getAttribute('fill') === '#107c41');
      const blackPath = paths.find((path) => path.getAttribute('fill') === '#1f1f1f');

      expect(guidePath).toBeTruthy();
      expect(blackPath).toBeTruthy();
      expect(paths.some((path) => path.getAttribute('fill') === '#8a8f98')).toBe(true);
      expect(guidePath?.getAttribute('d')).toMatch(/v3H4\.2/);
      expect(blackPath?.getAttribute('d')).toContain('v2.8');
    }

    for (const icon of ['alignLeft', 'alignCenter', 'alignRight']) {
      const svg = createIcon()(icon);
      const paths = Array.from(svg?.querySelectorAll('path') ?? []);
      const guidePath = paths.find((path) => path.getAttribute('fill') === '#107c41');
      const textRows = paths.find((path) => path.getAttribute('fill') === '#1f1f1f');

      expect(guidePath).toBeTruthy();
      expect(textRows).toBeTruthy();
      expect(paths).toHaveLength(2);
      expect(guidePath?.getAttribute('d')).toMatch(/h3v15\.4/);
      expect(textRows?.getAttribute('d')).toContain('v2.8');
      expect(textRows?.getAttribute('stroke')).toBeNull();
    }
  });

  it('uses Excel-style dedicated curved arrow icons for undo and redo', () => {
    for (const icon of ['undo', 'redo']) {
      const svg = createIcon()(icon);
      const paths = Array.from(svg?.querySelectorAll('path') ?? []);
      const greenSegments = paths.filter((path) => path.getAttribute('stroke') === '#107c41');
      const blackCurve = paths.find((path) => path.getAttribute('stroke') === '#1f1f1f');

      expect(svg?.getAttribute('fill')).toBeNull();
      expect(paths).toHaveLength(3);
      expect(greenSegments).toHaveLength(2);
      expect(blackCurve?.getAttribute('fill')).toBe('none');
      expect(blackCurve?.getAttribute('stroke-linecap')).toBe('round');
    }
  });

  it('falls back to currentColor Fluent SVGs for non-overridden icons', () => {
    const svg = createIcon()('add');

    expect(svg?.getAttribute('fill')).toBe('currentColor');
    expect(svg?.querySelectorAll('path')).toHaveLength(1);
  });

  it('returns null for unknown icon names', () => {
    expect(createIcon()('not-a-ribbon-icon')).toBeNull();
  });
});
