import { normalizeFormatLocale } from './locale.js';

const DEFAULT_CACHE_SIZE = 128;

function cacheKey(locale: string, options: object): string {
  const entries = Object.entries(options)
    .filter(([, value]) => value !== undefined)
    .sort(([left], [right]) => left.localeCompare(right));
  return JSON.stringify([locale, entries]);
}

/** Small LRU for the expensive formatter objects used repeatedly by cell paint. */
export class IntlFormatterCache {
  private readonly numberFormats = new Map<string, Intl.NumberFormat>();

  private readonly dateTimeFormats = new Map<string, Intl.DateTimeFormat>();

  constructor(private readonly maxEntries = DEFAULT_CACHE_SIZE) {}

  numberFormat(locale: string, options: Intl.NumberFormatOptions): Intl.NumberFormat {
    const normalizedLocale = normalizeFormatLocale(locale);
    const key = cacheKey(normalizedLocale, options);
    return this.getOrCreate(
      this.numberFormats,
      key,
      () => new Intl.NumberFormat(normalizedLocale, options),
    );
  }

  dateTimeFormat(locale: string, options: Intl.DateTimeFormatOptions): Intl.DateTimeFormat {
    const normalizedLocale = normalizeFormatLocale(locale);
    const key = cacheKey(normalizedLocale, options);
    return this.getOrCreate(
      this.dateTimeFormats,
      key,
      () => new Intl.DateTimeFormat(normalizedLocale, options),
    );
  }

  private getOrCreate<T>(cache: Map<string, T>, key: string, create: () => T): T {
    const cached = cache.get(key);
    if (cached !== undefined) {
      cache.delete(key);
      cache.set(key, cached);
      return cached;
    }

    const formatter = create();
    if (this.maxEntries <= 0) return formatter;
    cache.set(key, formatter);
    if (cache.size > this.maxEntries) {
      const oldest = cache.keys().next().value;
      if (oldest !== undefined) cache.delete(oldest);
    }
    return formatter;
  }
}

/** Separate bounded caches keep option sets from the independent formatters apart. */
export const intlFormatterCache = new IntlFormatterCache();
