import type { EngineCapabilities, Workbook } from './types.js';

/**
 * Probe the WASM module for optional bindings. As the engine grows,
 * methods are added one bundle at a time; this probe checks for each
 * method by name and flips the corresponding capability flag on iff
 * every method that flag depends on is present.
 *
 * The check is `typeof wb.<method> === 'function'`. Probing must be
 * free of side effects — the methods themselves are not invoked.
 */
export function detectCapabilities(wb: Workbook): EngineCapabilities {
  const w = wb as unknown as Record<string, unknown>;
  const has = (k: string): boolean => typeof w[k] === 'function';
  const all = (...keys: string[]): boolean => keys.every(has);

  return Object.freeze({
    merges: all('addMerge', 'getMerges', 'removeMerge', 'clearMerges'),
    cellFormatting: all(
      'getCellXfIndex',
      'setCellXfIndex',
      'getCellXf',
      'addFont',
      'addFill',
      'addBorder',
      'addNumFmt',
      'addXf',
      'getFont',
      'getFill',
      'getBorder',
      'getNumFmt',
    ),
    conditionalFormat: has('evaluateCfRange'),
    dataValidation: all('getValidations', 'addValidation', 'clearValidations'),
    sheetMutate: all('renameSheet', 'removeSheet', 'moveSheet'),
    insertDeleteRowsCols: all('insertRows', 'deleteRows', 'insertCols', 'deleteCols'),
    hiddenRowsCols: all('setRowHidden', 'setColumnHidden'),
    colRowSize: all('setColumnWidth', 'setRowHeight'),
    freeze: has('setSheetFreeze'),
    sheetZoom: has('setSheetZoom'),
    sheetTabHidden: has('setSheetTabHidden'),
    outlines: all('setColumnOutline', 'setRowOutline'),
    comments: all('getComment', 'setComment'),
    hyperlinks: all('getHyperlinks', 'addHyperlink', 'clearHyperlinks'),
    definedNameMutate: has('setDefinedName'),
    partialRecalc: has('partialRecalc'),
    iterativeProgress: has('setIterativeProgress'),
  });
}
