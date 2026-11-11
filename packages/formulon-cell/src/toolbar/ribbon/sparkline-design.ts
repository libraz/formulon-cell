import { SparklineType } from '@libraz/formulon';
import { isSheetProtected } from '../../commands/protection.js';
import {
  activeSparklineGroup,
  SPARKLINE_SHOW_COMMANDS,
  SPARKLINE_TYPE_COMMANDS,
} from '../../commands/sparkline-design.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import { projectDisabledState } from '../menu-a11y.js';
import { SPARKLINE_COMMANDS } from './sparkline-tabs.js';

/** Marks a button this projection disabled, so a disabled state set elsewhere is left alone. */
const OWN_DISABLED = 'sparklineDisabled';
/** The command's own tooltip, kept so the disabled reason can come and go beneath it. */
const OWN_TITLE = 'sparklineTitle';

/** Sync the Sparkline tab under `host` with the group of the active cell: the pressed Type and
 *  Show toggles, Markers (line sparklines only), and the greying a protected sheet asks for. The
 *  Axis menu stays enabled on a protected sheet, with every entry inert. The Style strip is
 *  refreshed to the group's preset. */
export const projectSparklineDesign = (host: HTMLElement, instance: SpreadsheetInstance): void => {
  const state = instance.store.getState();
  const group = activeSparklineGroup(state);
  const protectedSheet = isSheetProtected(state, state.selection.active.sheet);
  const settings = group?.settings;
  const reason = instance.i18n.strings.mac.sparkline.protectedLocation;
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
    const locked = id !== 'sparklineAxis' && protectedSheet;
    const disabled =
      !settings ||
      locked ||
      (id === 'sparklineShowMarkers' && settings.type !== SparklineType.Line);
    if (disabled) {
      if (!button.disabled || button.dataset[OWN_DISABLED]) {
        button.dataset[OWN_TITLE] ??= button.title;
        projectDisabledState(button, true, locked ? reason : null, {
          titlePrefix: button.dataset[OWN_TITLE],
        });
        button.dataset[OWN_DISABLED] = 'true';
      }
    } else if (button.dataset[OWN_DISABLED]) {
      projectDisabledState(button, false, null, { titlePrefix: button.dataset[OWN_TITLE] });
      delete button.dataset[OWN_DISABLED];
      delete button.dataset[OWN_TITLE];
    }
  }
  host
    .querySelector(`[data-ribbon-gallery="${SPARKLINE_COMMANDS.styleGallery}"]`)
    ?.dispatchEvent(new Event('fc-ribbon-control-refresh'));
};
