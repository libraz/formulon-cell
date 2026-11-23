import { isMacHost } from '../../extensions/ui-options.js';
import type { Strings } from '../../i18n/strings.js';
import { showMessage } from './prompt.js';

/** A command that refused to run: its range was too large, the engine could not carry it out, its host feature is off, or
 *  the selected picture or shape cannot be copied (one fixed in place, a locked shape, or a shape
 *  that keeps a paint as saved). */
export type CommandRefusalReason =
  | 'too-large'
  | 'engine-failed'
  | 'feature-disabled'
  | 'drawing-fixed-position'
  | 'shape-locked'
  | 'shape-paint-kept';

/** The message for `reason`, for an alert or an inline error. */
export const commandRefusalMessage = (reason: CommandRefusalReason, strings: Strings): string => {
  switch (reason) {
    case 'too-large':
      return strings.commandRefusal.tooLarge;
    case 'engine-failed':
      return strings.commandRefusal.engineFailed;
    case 'feature-disabled':
      return strings.commandRefusal.featureDisabled;
    case 'shape-locked':
      return strings.workbookObjects.shapeLocked;
    case 'shape-paint-kept':
      return strings.workbookObjects.shapePaintKept;
    case 'drawing-fixed-position':
      return strings.workbookObjects.drawingFixedPosition;
  }
};

/** Tell the user a command refused to run: on the Mac the title-less alert with one wide OK, on the
 *  default chrome an alert titled with the command's dialog. */
export const showCommandRefusal = (opts: {
  host: HTMLElement;
  strings: Strings;
  /** The command's dialog title. */
  title: string;
  reason: CommandRefusalReason;
}): Promise<void> =>
  showMessage({
    host: opts.host,
    title: opts.title,
    message: commandRefusalMessage(opts.reason, opts.strings),
    okLabel: opts.strings.pageSetup.ok,
    alert: isMacHost(opts.host),
  });
