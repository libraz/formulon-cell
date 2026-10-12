import type { ThemeResult } from '@libraz/formulon';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import type { OperationIntent } from '../commands/interaction-policy.js';
import { hydrateCellFormatsFromEngine } from '../engine/cell-format-sync.js';
import {
  WORKBOOK_THEME_PRESETS,
  type WorkbookThemePresetId,
} from '../engine/workbook-theme-presets.js';
import type { SpreadsheetInstance } from './types.js';

class WorkbookThemeRollbackError extends Error {}

export function applyWorkbookTheme(
  instance: SpreadsheetInstance,
  presetId: WorkbookThemePresetId,
  part: 'all' | 'colors' | 'fonts' = 'all',
  commandId = 'workbookTheme',
): boolean {
  const wb = instance.workbook;
  const intent: OperationIntent = {
    operation: 'format',
    origin: 'ribbon',
    commandId,
    effects: [{ kind: 'workbook' }],
  };
  const controller = interactionControllerFor(instance.store);
  if (controller && !controller.canExecute(intent).allowed) return false;
  if (
    !wb.capabilities.workbookTheme ||
    instance.store.getState().protection.protectedSheets.size > 0
  )
    return false;
  const before = wb.getTheme();
  if (!before || before.source === 2) return false;
  const preset = WORKBOOK_THEME_PRESETS[presetId];
  const after = {
    ...before,
    colors: part === 'fonts' ? [...before.colors] : [...preset.colors],
    fonts: part === 'colors' ? { ...before.fonts } : { ...preset.fonts },
  };
  if (
    JSON.stringify(before.colors) === JSON.stringify(after.colors) &&
    JSON.stringify(before.fonts) === JSON.stringify(after.fonts)
  )
    return true;
  const write = (next: Pick<ThemeResult, 'colors' | 'fonts'>): void => {
    const current = wb.getTheme();
    if (!current) throw new Error('Workbook theme is unavailable');
    const originalFormat = instance.store.getState().format;
    try {
      if (!wb.setThemeColors(next.colors)) throw new Error('Workbook theme colour update failed');
      if (!wb.setThemeFonts(next.fonts)) throw new Error('Workbook theme font update failed');
      wb.withEngineSyncMuted(() => {
        hydrateCellFormatsFromEngine(wb, instance.store, instance.store.getState().data.sheetIndex);
        instance.store.setState((state) => ({ ...state }));
      });
    } catch (error) {
      // Native setters materialize a theme part when the source was Default.
      // They can restore its colours/fonts, but cannot remove that new part.
      const colorsRestored = wb.setThemeColors(current.colors);
      const fontsRestored = wb.setThemeFonts(current.fonts);
      try {
        wb.withEngineSyncMuted(() => {
          instance.store.setState((state) => ({ ...state, format: originalFormat }));
        });
      } catch {
        // Store observers can throw after state has already been restored.
      }
      if (!colorsRestored || !fontsRestored)
        throw new WorkbookThemeRollbackError('Workbook theme update and rollback failed');
      throw error;
    }
  };
  try {
    write(after);
  } catch (error) {
    if (error instanceof WorkbookThemeRollbackError) throw error;
    return false;
  }
  instance.history.push({
    intent: { ...intent, origin: 'redo' },
    inverseIntent: { ...intent, origin: 'undo' },
    undo: () => write(before),
    redo: () => write(after),
  });
  return true;
}
