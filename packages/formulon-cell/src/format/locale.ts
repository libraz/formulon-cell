/** Canonical tags used for cell formatting and built-in language profiles. */
export const FORMAT_LOCALE_TAGS = {
  en: 'en-US',
  ja: 'ja-JP',
  de: 'de-DE',
  fr: 'fr-FR',
  zh: 'zh-CN',
  ko: 'ko-KR',
  th: 'th-TH',
} as const;

export const normalizeFormatLocale = (locale: string): string => {
  const normalized = locale.trim().replaceAll('_', '-');
  if (Object.hasOwn(FORMAT_LOCALE_TAGS, normalized.toLowerCase())) {
    return FORMAT_LOCALE_TAGS[normalized.toLowerCase() as keyof typeof FORMAT_LOCALE_TAGS];
  }
  try {
    return Intl.getCanonicalLocales(normalized)[0] ?? 'en-US';
  } catch {
    return 'en-US';
  }
};

/** Currency for the seven built-in locales; regional tags retain their country. */
export const currencySymbolForLocale = (locale: string): string => {
  const tag = normalizeFormatLocale(locale);
  const region = new Intl.Locale(tag).maximize().region;
  return (
    (
      {
        US: '$',
        JP: '¥',
        DE: '€',
        FR: '€',
        AT: '€',
        BE: '€',
        LU: '€',
        IE: '€',
        IT: '€',
        ES: '€',
        PT: '€',
        NL: '€',
        FI: '€',
        GR: '€',
        CA: 'C$',
        CH: 'CHF',
        CN: '¥',
        KR: '₩',
        TH: '฿',
        GB: '£',
      } as Record<string, string>
    )[region ?? ''] ?? '$'
  );
};

export const datePatternsForLocale = (
  locale: string,
): { short: string; long: string; time: string } => {
  const language = normalizeFormatLocale(locale).split('-')[0];
  switch (language) {
    case 'ja':
      return { short: 'yyyy/m/d', long: 'yyyy"年"m"月"d"日', time: 'H:MM' };
    case 'de':
      return { short: 'dd.mm.yyyy', long: 'dddd, d. mmmm yyyy', time: 'HH:MM' };
    case 'fr':
      return { short: 'dd/mm/yyyy', long: 'dddd d mmmm yyyy', time: 'HH:MM' };
    case 'zh':
      return { short: 'yyyy/m/d', long: 'yyyy"年"m"月"d"日', time: 'HH:MM' };
    case 'ko':
      return { short: 'yyyy-mm-dd', long: 'yyyy"년" m"월" d"일"', time: 'HH:MM' };
    case 'th':
      return { short: 'd/m/yyyy', long: 'd mmmm yyyy', time: 'HH:MM' };
    default:
      return { short: 'm/d/yyyy', long: 'mmmm d, yyyy', time: 'h:MM AM/PM' };
  }
};
