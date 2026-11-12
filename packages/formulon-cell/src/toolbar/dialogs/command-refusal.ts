import { isMacHost } from '../../extensions/ui-options.js';
import type { Strings } from '../../i18n/strings.js';
import { showMessage } from './prompt.js';

/** A command that refused to run: its range was too large, or the engine could not carry it out. */
export type CommandRefusalReason = 'too-large' | 'engine-failed';

/** The message for `reason`, for an alert or an inline error. */
export const commandRefusalMessage = (
  reason: CommandRefusalReason,
  strings: Strings['commandRefusal'],
): string => (reason === 'too-large' ? strings.tooLarge : strings.engineFailed);

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
    message: commandRefusalMessage(opts.reason, opts.strings.commandRefusal),
    okLabel: opts.strings.pageSetup.ok,
    alert: isMacHost(opts.host),
  });
