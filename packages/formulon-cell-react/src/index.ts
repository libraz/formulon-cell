export { Spreadsheet } from './Spreadsheet.js';
export type { SpreadsheetProps, SpreadsheetRef } from './Spreadsheet.js';

export { useI18n, useSelection, useSpreadsheet, useSpreadsheetEvent } from './hooks.js';

// Re-export the core types so consumers don't need to depend on the core
// package directly for typing — only at runtime.
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
