import { dictionaries, type Locale, type Strings } from '../../i18n/strings.js';
import { shouldShowFontOption } from './font-availability.js';

/**
 * A font option keeps the name written to a workbook separate from the text
 * shown in a picker.  In particular, Excel exposes the Japanese body theme
 * font as `游ゴシック Regular (本文)` while formulon stores `游ゴシック`.
 */
export interface FontFamilyOption {
  /** Exact value written to CellFormat / formulon. */
  value: string;
  /** Text shown before an optional localized role suffix. */
  label: string;
  section: 'theme' | 'all';
  role?: 'heading' | 'body';
  /** Alternate family names used only by the availability probe. */
  probeNames?: readonly string[];
  alwaysVisible?: boolean;
}

/** Input shape used when a caller already has a localized command model. */
export type FontFamilyOptionSource = Pick<FontFamilyOption, 'value' | 'label'>;
export type FontFamilyOptionFilter = (value: string, current: string, locale: Locale) => boolean;

const JAPANESE_BODY_VALUE = '游ゴシック';
const JAPANESE_BODY_DISPLAY_VALUE = '游ゴシック Regular';
const JAPANESE_HEADING_VALUE = '游ゴシック Light';

const canonicalValue = (raw: string): string => {
  switch (raw) {
    case JAPANESE_BODY_DISPLAY_VALUE:
      return JAPANESE_BODY_VALUE;
    case 'MS Pゴシック':
      return 'ＭＳ Ｐゴシック';
    case 'MS P明朝':
      return 'ＭＳ Ｐ明朝';
    default:
      return raw;
  }
};

const optionMetadata = (locale: Locale, value: string): Partial<FontFamilyOption> => {
  if (locale === 'ja' && value === JAPANESE_HEADING_VALUE) {
    return {
      section: 'theme',
      role: 'heading',
      alwaysVisible: true,
    };
  }
  if (locale === 'ja' && value === JAPANESE_BODY_VALUE) {
    return {
      section: 'theme',
      role: 'body',
      alwaysVisible: true,
      probeNames: [JAPANESE_BODY_VALUE, 'Yu Gothic', '游ゴシック体', 'YuGo-Medium'],
    };
  }
  if (locale !== 'ja' && value === 'Aptos Display') {
    return { section: 'theme', role: 'heading', alwaysVisible: true };
  }
  if (locale !== 'ja' && value === 'Aptos Narrow') {
    return { section: 'theme', role: 'body', alwaysVisible: true };
  }
  if (locale !== 'ja' && value === 'Aptos') {
    return { section: 'theme', alwaysVisible: true };
  }
  return { section: 'all' };
};

const labelForValue = (value: string): string =>
  value === JAPANESE_BODY_VALUE ? JAPANESE_BODY_DISPLAY_VALUE : value;

const normalizedValue = (value: string): string => value.trim().normalize('NFC');

const localeForStrings = (strings: Strings): Locale => {
  for (const [locale, dictionary] of Object.entries(dictionaries) as [Locale, Strings][]) {
    if (dictionary === strings) return locale;
  }
  return strings.ribbon.fontFamilies.some((family) => /[\u3000-\u9fff]/.test(family)) ? 'ja' : 'en';
};

const makeCatalog = (
  locale: Locale,
  sourceOptions: readonly FontFamilyOptionSource[],
): FontFamilyOption[] => {
  const seen = new Set<string>();
  const options: FontFamilyOption[] = [];
  for (const source of sourceOptions) {
    const value = canonicalValue(source.value);
    const key = normalizedValue(value);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    const metadata = optionMetadata(locale, value);
    options.push({
      value,
      label: source.label || labelForValue(value),
      section: metadata.section ?? 'all',
      role: metadata.role,
      probeNames: metadata.probeNames,
      alwaysVisible: metadata.alwaysVisible,
    });
  }
  return options;
};

/** Build the shared unfiltered catalog from a locale's ribbon dictionary. */
export const fontFamilyCatalog = (
  locale: Locale,
  sourceFamilies: readonly string[] = dictionaries[locale].ribbon.fontFamilies,
): readonly FontFamilyOption[] =>
  makeCatalog(
    locale,
    sourceFamilies.map((family) => ({
      value: family,
      label: labelForValue(canonicalValue(family)),
    })),
  );

/** Build a catalog from an already-localized command's value/label pairs. */
export const fontFamilyCatalogFromOptions = (
  locale: Locale,
  sourceOptions: readonly FontFamilyOptionSource[],
): readonly FontFamilyOption[] => makeCatalog(locale, sourceOptions);

/** Build the same filtered candidates used by both font pickers. */
export const fontFamilyOptions = (
  locale: Locale,
  current = '',
  sourceOptions?: readonly FontFamilyOptionSource[],
  filter?: FontFamilyOptionFilter,
): readonly FontFamilyOption[] => {
  const catalog = sourceOptions
    ? fontFamilyCatalogFromOptions(locale, sourceOptions)
    : fontFamilyCatalog(locale);
  const available = catalog
    .filter(
      (option) =>
        option.value === current ||
        (filter
          ? filter(option.value, current, locale)
          : shouldShowFontOption(option.value, current, locale, option)),
    )
    .map((option) =>
      current === JAPANESE_BODY_DISPLAY_VALUE && option.value === JAPANESE_BODY_VALUE
        ? { ...option, label: JAPANESE_BODY_VALUE }
        : option,
    );
  const currentKey = normalizedValue(current);
  if (currentKey && !available.some((option) => normalizedValue(option.value) === currentKey)) {
    // Imported values are intentionally retained verbatim.  This also keeps
    // the legacy `游ゴシック Regular` value visible without silently rewriting it.
    available.push({
      value: current,
      label: current,
      section: 'all',
      alwaysVisible: true,
    });
  }
  const theme = available.filter((option) => option.section === 'theme');
  const all = available
    .filter((option) => option.section === 'all')
    .sort((left, right) => {
      try {
        return new Intl.Collator(locale, { numeric: true, sensitivity: 'base' }).compare(
          left.label,
          right.label,
        );
      } catch {
        return left.label.localeCompare(right.label);
      }
    });
  return [...theme, ...all];
};

export const fontFamilyLabel = (locale: Locale, value: string): string => {
  const option = fontFamilyOptions(locale, value).find(
    (candidate) => normalizedValue(candidate.value) === normalizedValue(value),
  );
  return option?.label ?? value;
};

/** Map a known display label back to its persisted family value. */
export const fontFamilyValueFromLabel = (locale: Locale, input: string): string => {
  // The display label is also a legacy serialized value.  Keep typed legacy
  // text intact; picker rows use an explicit option value event below.
  if (locale === 'ja' && input === JAPANESE_BODY_DISPLAY_VALUE) return input;
  const option = fontFamilyCatalog(locale).find(
    (candidate) => candidate.label === input || candidate.value === input,
  );
  return option?.value ?? input;
};

export const fontFamilyLocaleForStrings = localeForStrings;
