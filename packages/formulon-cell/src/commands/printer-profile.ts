import type { PageMargins, PageOrientation, PageSetup, PaperSize } from '../store/store.js';

export interface PrinterProfile {
  id?: string;
  name?: string;
  paperSize?: PaperSize;
  orientation?: PageOrientation;
  printableBounds: Partial<PageMargins>;
}

export interface HostPrinterPaperOption {
  id?: string;
  label?: string;
  paperSize?: PaperSize | string;
  orientation?: PageOrientation | string;
  printableBounds?: Partial<PageMargins>;
  hardwareMarginsInches?: Partial<PageMargins>;
}

export interface HostPrinterDevice {
  id?: string;
  name?: string;
  paperOptions?: readonly HostPrinterPaperOption[];
  paperSize?: PaperSize | string;
  orientation?: PageOrientation | string;
  printableBounds?: Partial<PageMargins>;
  hardwareMarginsInches?: Partial<PageMargins>;
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

const hostPaperSize = (value: string | undefined): PaperSize | undefined =>
  PAPER_SIZES.has(value as PaperSize) ? (value as PaperSize) : undefined;

const hostOrientation = (value: string | undefined): PageOrientation | undefined =>
  ORIENTATIONS.has(value as PageOrientation) ? (value as PageOrientation) : undefined;

const printerProfileName = (
  device: HostPrinterDevice,
  paper: HostPrinterPaperOption | HostPrinterDevice,
): string | undefined => {
  const deviceName = device.name?.trim();
  const paperLabel = 'label' in paper ? paper.label?.trim() : undefined;
  if (deviceName && paperLabel) return `${deviceName} - ${paperLabel}`;
  return deviceName || paperLabel || undefined;
};

export const printerProfilesFromHostDevices = (
  devices: readonly HostPrinterDevice[] | undefined,
): readonly PrinterProfile[] | undefined => {
  if (!devices) return undefined;
  const profiles: PrinterProfile[] = [];
  devices.forEach((device, deviceIndex) => {
    const papers = device.paperOptions?.length ? device.paperOptions : [device];
    papers.forEach((paper, paperIndex) => {
      const paperSize = hostPaperSize(paper.paperSize);
      const orientation = hostOrientation(paper.orientation);
      const deviceId =
        normalizePrinterProfileId(device.id) ?? normalizePrinterProfileId(device.name);
      const paperOption = paper === device ? undefined : (paper as HostPrinterPaperOption);
      const paperId =
        normalizePrinterProfileId(paperOption?.id) ?? normalizePrinterProfileId(paperOption?.label);
      const fallbackId = [deviceIndex, paperIndex, paperSize, orientation]
        .filter(Boolean)
        .join(':');
      profiles.push({
        id: [deviceId, paperId ?? fallbackId].filter(Boolean).join(':') || undefined,
        name: printerProfileName(device, paper),
        paperSize,
        orientation,
        printableBounds:
          paper.hardwareMarginsInches ??
          paper.printableBounds ??
          device.hardwareMarginsInches ??
          device.printableBounds ??
          {},
      });
    });
  });
  return normalizePrinterProfiles(profiles);
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
