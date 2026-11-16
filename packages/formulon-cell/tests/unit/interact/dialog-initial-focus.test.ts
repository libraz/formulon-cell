import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  disposeMacConsolidateDialog,
  openMacConsolidateDialog,
} from '../../../src/interact/mac-consolidate-dialog.js';
import {
  disposeMacGoalSeekDialog,
  openMacGoalSeekDialog,
} from '../../../src/interact/mac-goal-seek-dialog.js';
import {
  disposeMacSubtotalDialog,
  openMacSubtotalDialog,
} from '../../../src/interact/mac-subtotal-dialog.js';
import {
  byId,
  createMacDataDialogFixture,
  type MacDataDialogFixture,
} from './mac-data-dialog-fixture.js';

const nextFrame = (): Promise<void> => new Promise((r) => requestAnimationFrame(() => r()));

const cases: {
  name: string;
  open: (fx: MacDataDialogFixture) => void;
  dispose: (fx: MacDataDialogFixture) => void;
  /** An element the user moves to right after the dialog opens. */
  moveTo: string;
}[] = [
  {
    name: 'Subtotal',
    open: (fx) => openMacSubtotalDialog(fx.instance),
    dispose: (fx) => disposeMacSubtotalDialog(fx.instance),
    moveTo: 'fc-mac-subtotal-function',
  },
  {
    name: 'Consolidate',
    open: (fx) => openMacConsolidateDialog(fx.instance),
    dispose: (fx) => disposeMacConsolidateDialog(fx.instance),
    moveTo: 'fc-mac-consolidate-function',
  },
  {
    name: 'Goal Seek',
    open: (fx) => openMacGoalSeekDialog(fx.instance),
    dispose: (fx) => disposeMacGoalSeekDialog(fx.instance),
    moveTo: 'fc-mac-goalseek-target-value',
  },
];

describe('data dialogs initial focus', () => {
  let fx: MacDataDialogFixture;
  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
  });
  afterEach(() => fx.dispose());

  for (const c of cases) {
    it(`${c.name}: does not steal focus back after the user has moved`, async () => {
      c.open(fx);
      const target = byId<HTMLElement>(c.moveTo);
      target.focus();
      expect(document.activeElement).toBe(target);
      await nextFrame();
      await nextFrame();
      expect(document.activeElement).toBe(target);
      c.dispose(fx);
    });
  }
});
