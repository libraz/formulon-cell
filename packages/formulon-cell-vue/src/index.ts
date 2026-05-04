// Re-export the core types so consumers don't need a separate
// `@libraz/formulon-cell` type-only dependency.
export type {
  CellChangeEvent,
  CellRegistry,
  Extension,
  ExtensionContext,
  ExtensionHandle,
  ExtensionInput,
  FeatureFlags,
  FeatureId,
  FormulaRegistry,
  I18nController,
  LocaleChangeEvent,
  MountOptions,
  RecalcEvent,
  SelectionChangeEvent,
  SpreadsheetEventHandler,
  SpreadsheetEventName,
  SpreadsheetEvents,
  SpreadsheetInstance,
  ThemeChangeEvent,
  ThemeName,
  WorkbookChangeEvent,
  WorkbookHandle,
} from '@libraz/formulon-cell';
export { presets } from '@libraz/formulon-cell';

export {
  useI18n,
  useSelection,
  useSpreadsheet,
  useSpreadsheetEvent,
} from './composables.js';
export type { SpreadsheetExposed } from './Spreadsheet.js';
export { Spreadsheet } from './Spreadsheet.js';
