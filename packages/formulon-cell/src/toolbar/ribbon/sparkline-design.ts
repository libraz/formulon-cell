import { SparklineType } from '@libraz/formulon';
import { isSheetProtected } from '../../commands/protection.js';
import {
  activeSparklineGroup,
  SPARKLINE_SHOW_COMMANDS,
  SPARKLINE_TYPE_COMMANDS,
} from '../../commands/sparkline-design.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { projectDisabledState } from '../menu-a11y.js';

/** Marks a button this projection disabled, so a disabled state set elsewhere is left alone. */
const OWN_DISABLED = 'sparklineDisabled';

/** Sync the Sparkline tab under `host` with the group of the active cell: the pressed Type and
 *  Show toggles, Markers (line sparklines only), and the greying a protected sheet asks for. The
 *  Axis menu stays enabled on a protected sheet, with every entry inert. */
export const projectSparklineDesign = (host: HTMLElement, instance: SpreadsheetInstance): void => {
  const state = instance.store.getState();
  const group = activeSparklineGroup(state);
  const protectedSheet = isSheetProtected(state, state.selection.active.sheet);
  const settings = group?.settings;
  for (const button of host.querySelectorAll<HTMLButtonElement>('[data-ribbon-command]')) {
    const id = button.dataset.ribbonCommand ?? '';
    const type = SPARKLINE_TYPE_COMMANDS[id];
    const flag = SPARKLINE_SHOW_COMMANDS[id];
    const isSparkline = type !== undefined || flag !== undefined || id.startsWith('sparkline');
    if (!isSparkline) continue;
    if (type !== undefined || flag !== undefined) {
      const pressed =
        type !== undefined ? settings?.type === type : !!flag && settings?.[flag] === true;
      button.classList.toggle('fc-tb__rb--active', pressed);
      button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
    }
    const disabled =
      !settings ||
      (id !== 'sparklineAxis' && protectedSheet) ||
      (id === 'sparklineShowMarkers' && settings.type !== SparklineType.Line);
    if (disabled) {
      if (!button.disabled || button.dataset[OWN_DISABLED]) {
        projectDisabledState(button, true, null);
        button.dataset[OWN_DISABLED] = 'true';
      }
    } else if (button.dataset[OWN_DISABLED]) {
      projectDisabledState(button, false, null);
      delete button.dataset[OWN_DISABLED];
    }
  }
};
