import { describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { WORKBOOK_THEME_PRESETS } from '../../../src/engine/workbook-theme-presets.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { applyWorkbookTheme } from '../../../src/mount/workbook-theme-action.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('workbook theme actions', () => {
  it('edits colors and fonts independently, with native undo/redo and save/load', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const instance = { workbook, store, history } as SpreadsheetInstance;
    workbook.attachStore(store);
    try {
      const original = workbook.getTheme();
      expect(original).not.toBeNull();
      expect(applyWorkbookTheme(instance, 'gray', 'colors')).toBe(true);
      expect(workbook.getTheme()?.colors).toEqual(WORKBOOK_THEME_PRESETS.gray.colors);
      expect(workbook.getTheme()?.fonts).toEqual(original?.fonts);
      expect(store.getState().ui.theme).toBe('paper');
      history.undo();
      expect(workbook.getTheme()?.colors).toEqual(original?.colors);
      history.redo();
      expect(workbook.getTheme()?.colors).toEqual(WORKBOOK_THEME_PRESETS.gray.colors);
      expect(applyWorkbookTheme(instance, 'office', 'fonts')).toBe(true);
      expect(workbook.getTheme()?.colors).toEqual(WORKBOOK_THEME_PRESETS.gray.colors);
      expect(workbook.getTheme()?.fonts).toEqual(WORKBOOK_THEME_PRESETS.office.fonts);
      const loaded = await WorkbookHandle.loadBytes(workbook.save());
      try {
        expect(loaded.getTheme()?.colors).toEqual(WORKBOOK_THEME_PRESETS.gray.colors);
        expect(loaded.getTheme()?.fonts).toEqual(WORKBOOK_THEME_PRESETS.office.fonts);
      } finally {
        loaded.dispose();
      }
    } finally {
      workbook.dispose();
    }
  });

  it('authorizes themes as workbook-wide formatting', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    try {
      controller.setPolicy(viewerPolicy());
      const original = workbook.getTheme();
      expect(applyWorkbookTheme({ workbook, store, history } as SpreadsheetInstance, 'gray')).toBe(
        false,
      );
      expect(workbook.getTheme()).toEqual(original);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('restores the persisted theme when format hydration fails', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const original = workbook.getTheme();
    const format = store.getState().format;
    const failure = vi.spyOn(workbook, 'physicalCells').mockImplementation(() => {
      throw new Error('hydrate failed');
    });
    try {
      expect(applyWorkbookTheme({ workbook, store, history } as SpreadsheetInstance, 'gray')).toBe(
        false,
      );
      expect(workbook.getTheme()).toMatchObject({
        colors: original?.colors,
        fonts: original?.fonts,
      });
      expect(store.getState().format).toBe(format);
      expect(history.canUndo()).toBe(false);
    } finally {
      failure.mockRestore();
      workbook.dispose();
    }
  });

  it('keeps the leaf command and replay origin when reauthorizing history', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    const commandId = 'mac.page.color.gray';
    try {
      controller.setPolicy({
        operations: { format: true },
        restrict: (context) => context.commandId === commandId,
      });
      expect(
        applyWorkbookTheme(
          { workbook, store, history } as SpreadsheetInstance,
          'gray',
          'colors',
          commandId,
        ),
      ).toBe(true);
      const applied = workbook.getTheme();
      controller.setPolicy({
        operations: { format: true },
        restrict: (context) => context.origin !== 'undo',
      });
      expect(history.undo()).toBe(false);
      expect(workbook.getTheme()).toEqual(applied);
      controller.setPolicy({
        operations: { format: true },
        restrict: (context) => context.commandId === commandId,
      });
      expect(history.undo()).toBe(true);
      controller.setPolicy({
        operations: { format: true },
        restrict: (context) => context.origin !== 'redo',
      });
      expect(history.redo()).toBe(false);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });
});
