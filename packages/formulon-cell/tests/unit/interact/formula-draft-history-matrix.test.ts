import { describe, expect, it } from 'vitest';
import {
  FormulaDraftHistory,
  type FormulaDraftSnapshot,
} from '../../../src/interact/formula-draft-history.js';

// coverwise generate, strength=2, seed=42: all 44 parameter pairs covered.
const cases = [
  { transaction: 'group', mode: 'edit', direction: 'none', checkpoint: 'direct' },
  { transaction: 'composition', mode: 'point', direction: 'forward', checkpoint: 'suspended' },
  { transaction: 'command', mode: 'edit', direction: 'backward', checkpoint: 'suspended' },
  { transaction: 'input', mode: 'point', direction: 'none', checkpoint: 'suspended' },
  { transaction: 'group', mode: 'point', direction: 'backward', checkpoint: 'direct' },
  { transaction: 'composition', mode: 'edit', direction: 'none', checkpoint: 'suspended' },
  { transaction: 'group', mode: 'edit', direction: 'forward', checkpoint: 'suspended' },
  { transaction: 'input', mode: 'edit', direction: 'backward', checkpoint: 'direct' },
  { transaction: 'input', mode: 'point', direction: 'forward', checkpoint: 'direct' },
  { transaction: 'composition', mode: 'point', direction: 'backward', checkpoint: 'direct' },
  { transaction: 'command', mode: 'point', direction: 'none', checkpoint: 'direct' },
  { transaction: 'command', mode: 'edit', direction: 'forward', checkpoint: 'suspended' },
] as const;

describe('draft history snapshot combinations', () => {
  it.each(cases)(
    '$transaction $mode $direction $checkpoint restores the original selection and mode',
    ({ transaction, mode, direction, checkpoint }) => {
      const before: FormulaDraftSnapshot = { value: '=SU', start: 1, end: 3, direction, mode };
      const after: FormulaDraftSnapshot = {
        value: '=SUM(',
        start: 5,
        end: 5,
        direction: 'none',
        mode: transaction === 'group' || transaction === 'command' ? 'point' : mode,
      };
      const history = new FormulaDraftHistory(before);
      if (transaction === 'input') {
        history.beforeInput(before);
        history.input(after);
      } else if (transaction === 'command') {
        history.beginCommand(before);
        history.finishCommand(after);
      } else if (transaction === 'group') {
        history.beginGroup(before, 'reference');
        history.updateGroup({ ...after, value: '=SUM', start: 4, end: 4 }, 'reference');
        history.updateGroup(after, 'reference');
      } else {
        history.startComposition(before);
        history.input({ ...after, value: '=す', start: 2, end: 2 }, { composing: true });
        history.endComposition(after);
      }

      let restored = history;
      if (checkpoint === 'suspended') {
        const captured = history.captureState();
        // Later changes in another owner must not mutate the captured journal.
        history.recordCommand(after, { ...after, value: '=OTHER(', start: 7, end: 7 });
        restored = new FormulaDraftHistory();
        restored.restoreState(captured);
      }
      expect(restored.undo()).toEqual(before);
      expect(restored.undo()).toBeNull();
      expect(restored.redo()).toEqual(after);
      expect(restored.redo()).toBeNull();
    },
  );
});
