import type { Disposer } from '../api/events.js';
import type { DeepPartial, Strings } from './strings.js';

/** UI language and strings. Change the language with `update({ locale })`. */
export interface I18n {
  readonly locale: string;
  readonly strings: Strings;
  /** Overrides strings of `locale`; takes effect at once when it is the current locale. */
  override(locale: string, strings: DeepPartial<Strings>): void;
  on(name: 'change', listener: (event: { locale: string; strings: Strings }) => void): Disposer;
}
