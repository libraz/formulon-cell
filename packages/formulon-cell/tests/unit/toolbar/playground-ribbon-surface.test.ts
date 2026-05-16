import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildRibbonModel } from '../../../src/toolbar/ribbon-model.js';

const playgroundMainSource = (): string => {
  // The Cells menu DOM has been extracted to apps/playground/src/ribbon/menus/
  // factories (home.ts, insert.ts, ...). To keep these source-scrape checks
  // pointing at the playground surface, concatenate main.ts with every menu
  // factory module so the assertions match regardless of which file actually
  // owns a given menuButton/case.
  const roots = [resolve(process.cwd(), '../../'), resolve(process.cwd())];
  const playgroundRoot = roots.find((r) => existsSync(`${r}/apps/playground/src/main.ts`));
  expect(playgroundRoot).toBeTruthy();
  const root = playgroundRoot!;
  const files = [
    `${root}/apps/playground/src/main.ts`,
    `${root}/apps/playground/src/ribbon/menus/borders.ts`,
    `${root}/apps/playground/src/ribbon/menus/conditional.ts`,
    `${root}/apps/playground/src/ribbon/menus/general.ts`,
    `${root}/apps/playground/src/ribbon/menus/home.ts`,
    `${root}/apps/playground/src/ribbon/menus/insert.ts`,
    `${root}/apps/playground/src/ribbon/menus/page-layout.ts`,
    `${root}/apps/playground/src/ribbon/menus/formulas.ts`,
    `${root}/apps/playground/src/ribbon/menus/paste.ts`,
    `${root}/apps/playground/src/ribbon/menus/review.ts`,
    `${root}/apps/playground/src/ribbon/menus/styles.ts`,
    `${root}/apps/playground/src/ribbon/menus/text-orientation.ts`,
  ].filter((path) => existsSync(path));
  return files.map((path) => readFileSync(path, 'utf8')).join('\n');
};

const extractSwitchCases = (source: string, functionName: string): Set<string> => {
  const start = source.indexOf(`const ${functionName}`);
  expect(start).toBeGreaterThanOrEqual(0);
  const end = source.indexOf('\n};', start);
  expect(end).toBeGreaterThan(start);
  const body = source.slice(start, end);
  return new Set(Array.from(body.matchAll(/case '([^']+)'/g), (match) => match[1] ?? ''));
};

const extractDynamicDropdownCommands = (source: string): Set<string> => {
  const start = source.indexOf('const dynamicDropdownSpecForButton');
  expect(start).toBeGreaterThanOrEqual(0);
  const end = source.indexOf('\n};', start);
  expect(end).toBeGreaterThan(start);
  const body = source.slice(start, end);
  return new Set(Array.from(body.matchAll(/command === '([^']+)'/g), (match) => match[1] ?? ''));
};

const extractLegacyCommands = (source: string): Set<string> => {
  const match = source.match(/const legacyCommandIds:[\s\S]*?=\s*\{([\s\S]*?)\n\};/);
  expect(match?.[1]).toBeTruthy();
  return new Set(
    Array.from((match?.[1] ?? '').matchAll(/^\s*([A-Za-z0-9]+):/gm), (m) => m[1] ?? ''),
  );
};

describe('playground ribbon command surface', () => {
  it('routes every shared button command through a concrete playground handler', () => {
    const source = playgroundMainSource();
    const applyCases = extractSwitchCases(source, 'applyRibbonCommand');
    const dynamicCommands = extractDynamicDropdownCommands(source);
    const legacyCommands = extractLegacyCommands(source);
    const specialClickCommands = new Set(['borders', 'freeze', 'printArea', 'symbolInsert']);
    const covered = new Set([
      ...applyCases,
      ...dynamicCommands,
      ...legacyCommands,
      ...specialClickCommands,
    ]);

    const modelButtonIds = buildRibbonModel('en')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => !['break', 'color', 'select'].includes(command.kind ?? 'button'))
      .map((command) => command.id);

    const missing = modelButtonIds.filter((id) => !covered.has(id));
    expect(missing).toEqual([]);
  });

  it('uses dedicated routes for chart insertion and data validation', () => {
    const source = playgroundMainSource();
    expect(source).toContain("case 'chartInsert':");
    expect(source).toContain('createChartFromSelection();');
    expect(source).toContain("case 'dataValidation':");
    expect(source).toContain('i.openDataValidationDialog();');
    expect(source).not.toContain("case 'dataValidation':\n      i.openFormatDialog('more');");
  });

  it('records playground conditional-formatting menu changes as undoable operations', () => {
    const source = playgroundMainSource();
    expect(source).toContain('recordConditionalRulesChange,');
    expect(source).toContain('recordConditionalRulesChange(i.history, i.store, () => {');
    expect(source).toContain('mutators.addConditionalRule(i.store, rule);');
    expect(source).toContain('changed = applyConditionalPresetAction(i.store, action, range);');
    expect(source).not.toContain('if (applyConditionalPresetAction(i.store, action, range))');
  });

  it('keeps playground ribbon text paste undoable and selects the pasted range', () => {
    const source = playgroundMainSource();
    expect(source).toContain("if (action === 'all' || action === 'values') {");
    expect(source).toContain('let result: ReturnType<typeof pasteTSV> = null;');
    expect(source).toContain('result = pasteTSV(i.store.getState(), i.workbook, text);');
    expect(source).toContain('if (result) mutators.setRange(i.store, result.writtenRange);');
  });

  it('exposes concrete playground Cells insert/delete actions without relying only on prompts', () => {
    const source = playgroundMainSource();
    expect(source).toContain("menuButton(t.insertShiftDown, 'cellInsert', 'shift-down')");
    expect(source).toContain("menuButton(t.insertShiftRight, 'cellInsert', 'shift-right')");
    expect(source).toContain("menuButton(sheetTabs.insertSheet, 'cellInsert', 'sheet')");
    expect(source).toContain(
      "insertCells(i.store, i.workbook, i.history, range, action === 'shift-down' ? 'down' : 'right')",
    );
    expect(source).toContain('const added = addSheet(i.store, i.workbook, i.history);');

    expect(source).toContain("menuButton(t.deleteShiftUp, 'cellDelete', 'shift-up')");
    expect(source).toContain("menuButton(t.deleteShiftLeft, 'cellDelete', 'shift-left')");
    expect(source).toContain("menuButton(sheetTabs.deleteSheet, 'cellDelete', 'sheet')");
    expect(source).toContain(
      "deleteCells(i.store, i.workbook, i.history, range, action === 'shift-up' ? 'up' : 'left')",
    );
    expect(source).toContain('removeSheet(i.store, i.workbook, before)');
  });

  it('routes playground Cells > Format sheet actions and tab colors through concrete handlers', () => {
    const source = playgroundMainSource();
    expect(source).toContain("menuButton(sheetTabs.rename, 'cellFormat', 'rename-sheet')");
    expect(source).toContain("menuButton(sheetTabs.moveLeft, 'cellFormat', 'move-sheet-left')");
    expect(source).toContain("menuButton(sheetTabs.moveRight, 'cellFormat', 'move-sheet-right')");
    expect(source).toContain("menuButton(sheetTabs.hideSheet, 'cellFormat', 'hide-sheet')");
    expect(source).toContain("menuButton(sheetTabs.unhideSheet, 'cellFormat', 'unhide-sheet')");
    expect(source).toContain(
      "menuButton(`${sheetTabs.tabColor}: ${sheetTabs.noColor}`, 'cellFormat', 'tab-color-none')",
    );
    expect(source).toContain('renameSheet(i.workbook, sheet, name.trim(), i.store, i.history)');
    expect(source).toContain('moveSheet(i.store, i.workbook, sheet, target, i.history)');
    expect(source).toContain('setSheetHidden(i.store, i.workbook, i.history, sheet, true)');
    expect(source).toContain('setSheetHidden(i.store, i.workbook, i.history, hidden, false)');
    expect(source).toContain('mutators.setSheetTabColor(i.store, sheet, tabColor);');
  });
});
