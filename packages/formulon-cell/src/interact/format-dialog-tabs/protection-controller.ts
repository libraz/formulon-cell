// Protection tab behaviour for the Format Cells dialog: the locked and
// hidden-formula flags. The tab's DOM is built inline by
// `format-dialog-view.ts`.

import type { makeCheckbox } from '../format-dialog-dom.js';
import { type FormatTabContext, type FormatTabController, markMixedCheck } from './controller.js';

export interface ProtectionTabRefs {
  lockedCk: ReturnType<typeof makeCheckbox>;
  hiddenFormulaCk: ReturnType<typeof makeCheckbox>;
}

export function attachProtectionTab(
  ctx: FormatTabContext,
  refs: ProtectionTabRefs,
): FormatTabController {
  const { draft, on, touch, isMixed } = ctx;
  const { lockedCk, hiddenFormulaCk } = refs;

  const sync = (): void => {
    lockedCk.input.checked = draft.locked;
    hiddenFormulaCk.input.checked = draft.formulaHidden;
  };

  const syncMixed = (): void => {
    markMixedCheck(lockedCk.input, isMixed('locked'), 'locked');
    markMixedCheck(hiddenFormulaCk.input, isMixed('formulaHidden'), 'formulaHidden');
  };

  const onLockedChange = (): void => {
    touch('locked');
    draft.locked = lockedCk.input.checked;
  };
  const onHiddenFormulaChange = (): void => {
    touch('formulaHidden');
    draft.formulaHidden = hiddenFormulaCk.input.checked;
  };

  on(lockedCk.input, 'change', onLockedChange);
  on(hiddenFormulaCk.input, 'change', onHiddenFormulaChange);

  return { sync, syncMixed };
}
