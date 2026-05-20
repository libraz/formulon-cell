import type { PageMargins, PageOrientation, PageSetup, PaperSize } from '../store/store.js';

export interface PrinterProfile {
  id?: string;
  name?: string;
  paperSize?: PaperSize;
  orientation?: PageOrientation;
  printableBounds: Partial<PageMargins>;
}

const PAPER_SIZES = new Set<PaperSize>(['A4', 'A3', 'A5', 'letter', 'legal', 'tabloid']);
const ORIENTATIONS = new Set<PageOrientation>(['portrait', 'landscape']);

export const normalizePrinterProfileId = (id: string | undefined): string | undefined => {
  const normalized = id?.trim();
  return normalized ? normalized : undefined;
};

export const normalizePrintableBounds = (
  bounds: Partial<PageMargins> | undefined,
): PageMargins | undefined => {
  if (!bounds) return undefined;
  const read = (value: number | undefined): number =>
    Number.isFinite(value) ? Math.max(0, value ?? 0) : 0;
  return {
    top: read(bounds.top),
    right: read(bounds.right),
    bottom: read(bounds.bottom),
    left: read(bounds.left),
  };
};

export const normalizePrinterProfile = (profile: PrinterProfile): PrinterProfile => {
  const id = profile.id?.trim();
  const name = profile.name?.trim();
  return {
    ...(id ? { id } : {}),
    ...(name ? { name } : {}),
    ...(profile.paperSize && PAPER_SIZES.has(profile.paperSize)
      ? { paperSize: profile.paperSize }
      : {}),
    ...(profile.orientation && ORIENTATIONS.has(profile.orientation)
      ? { orientation: profile.orientation }
      : {}),
    printableBounds: normalizePrintableBounds(profile.printableBounds) ?? {
      top: 0,
      right: 0,
      bottom: 0,
      left: 0,
    },
  };
};

export const normalizePrinterProfiles = (
  profiles: readonly PrinterProfile[] | undefined,
): readonly PrinterProfile[] | undefined => {
  if (!profiles) return undefined;
  const seen = new Set<string>();
  const normalized: PrinterProfile[] = [];
  for (const profile of profiles.map(normalizePrinterProfile)) {
    const key = profile.id
      ? `id:${profile.id}`
      : [
          profile.name ? `name:${profile.name}` : '',
          profile.paperSize ?? '',
          profile.orientation ?? '',
        ]
          .filter(Boolean)
          .join('|');
    if (key && seen.has(key)) continue;
    if (key) seen.add(key);
    normalized.push(profile);
  }
  return normalized;
};

export function resolvePrinterProfileBounds(
  setup: PageSetup,
  profiles: readonly PrinterProfile[],
  preferredId?: string,
): PageMargins | undefined {
  const normalizedPreferredId = normalizePrinterProfileId(preferredId);
  const candidates = profiles.filter(
    (profile) =>
      (!profile.paperSize || profile.paperSize === setup.paperSize) &&
      (!profile.orientation || profile.orientation === setup.orientation),
  );
  const preferred = normalizedPreferredId
    ? (candidates.find((profile) => profile.id === normalizedPreferredId) ??
      profiles.find((profile) => profile.id === normalizedPreferredId))
    : undefined;
  const best =
    preferred ??
    candidates.find((profile) => profile.paperSize && profile.orientation) ??
    candidates.find((profile) => profile.paperSize) ??
    candidates.find((profile) => profile.orientation) ??
    candidates[0];
  return normalizePrintableBounds(best?.printableBounds);
}
