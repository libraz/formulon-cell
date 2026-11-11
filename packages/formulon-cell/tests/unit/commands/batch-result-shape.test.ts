import { describe, expectTypeOf, it } from 'vitest';
import type { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import type { pasteHtml } from '../../../src/commands/clipboard/html.js';
import type {
  insertCopiedBand,
  insertCopiedCellsFromTSV,
} from '../../../src/commands/clipboard/insert-copied-cells.js';
import type { pasteTSV } from '../../../src/commands/clipboard/paste.js';
import type {
  insertCopiedBandAction,
  insertCopiedCellsAction,
  pasteSpecialAction,
} from '../../../src/commands/clipboard/paste-actions.js';
import type { pasteLink, pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import type {
  executeFillHandle,
  executeRibbonFillAction,
  justifySelection,
} from '../../../src/commands/fill.js';
import type { importCSV } from '../../../src/commands/import-export.js';
import type { ChangeBatchResult } from '../../../src/commands/interaction-policy.js';
import type { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import type { applyPasteSpecialFrom } from '../../../src/interact/dialog-entry.js';

// Every command that writes a batch of cells answers with the one batch result, refusals
// included, so a new batch command cannot come back returning null or a boolean.
describe('batch-write commands', () => {
  it('return the one batch result', () => {
    expectTypeOf<ReturnType<typeof pasteTSV>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof pasteHtml>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof pasteSpecial>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof pasteLink>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof pasteSpecialAction>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof applyPasteSpecialFrom>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof insertCopiedBand>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof insertCopiedCellsFromTSV>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof insertCopiedBandAction>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof insertCopiedCellsAction>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof executeRibbonFillAction>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof justifySelection>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof executeFillHandle>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof executeRibbonClearAction>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof clearSelectedContents>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<typeof importCSV>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<keyof ChangeBatchResult>().toEqualTypeOf<
      'status' | 'applied' | 'rejections' | 'range'
    >();
  });
});
