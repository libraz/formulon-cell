import type { ToolbarLang } from '../../ribbon-model.js';
import { MAC_BUNDLE_LABELS } from './labels-bundle.js';
import type { Label } from './labels-data.js';

/** The generated Mac label of `key` in every locale, or undefined when none was generated. */
export const findMacBundleLabel = (key: string): Label | undefined => MAC_BUNDLE_LABELS[key];

/** The generated Mac label of `key`; a key the generator did not emit is a programming error. */
export const macBundleLabel = (key: string): Label => {
  const label = MAC_BUNDLE_LABELS[key];
  if (!label) throw new Error(`Missing generated Mac label: ${key}`);
  return label;
};

/** The generated Mac text of `key` in `lang`. */
export const macBundleText = (key: string, lang: ToolbarLang): string => macBundleLabel(key)[lang];
