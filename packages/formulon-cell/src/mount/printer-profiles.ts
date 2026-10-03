import {
  normalizePrinterProfileId,
  normalizePrinterProfiles,
  type PrinterProfile,
  resolvePrinterProfileBounds,
} from '../commands/printer-profile.js';
import type { PageMargins, PageSetup } from '../store/store.js';
import type { MountOptions } from './types.js';

export interface PrinterProfilesDeps {
  profiles: MountOptions['printerProfiles'];
  profileId: MountOptions['printerProfileId'];
  refreshHook: MountOptions['refreshPrinterProfiles'];
}

export interface PrinterProfilesState {
  getProfiles: () => readonly PrinterProfile[] | undefined;
  getProfileId: () => string | undefined;
  setProfiles: (next: readonly PrinterProfile[] | undefined) => void;
  setProfileId: (next: string | undefined) => void;
  refresh: () => Promise<readonly PrinterProfile[] | undefined>;
  /** Printable bounds for `setup`; `undefined` when no profiles are configured. */
  printableBounds: (
    setup: PageSetup,
    selectedProfileId: string | undefined,
  ) => Partial<PageMargins> | null | undefined;
}

export function createPrinterProfiles(deps: PrinterProfilesDeps): PrinterProfilesState {
  let printerProfiles = normalizePrinterProfiles(deps.profiles);
  let printerProfileId = normalizePrinterProfileId(deps.profileId);
  const refreshHook = deps.refreshHook;
  return {
    getProfiles: () => printerProfiles,
    getProfileId: () => printerProfileId,
    setProfiles(next) {
      printerProfiles = normalizePrinterProfiles(next);
    },
    setProfileId(next) {
      printerProfileId = normalizePrinterProfileId(next);
    },
    async refresh() {
      const next = await refreshHook?.();
      if (next !== undefined) printerProfiles = normalizePrinterProfiles(next);
      return printerProfiles;
    },
    printableBounds: (setup, selectedProfileId) =>
      printerProfiles
        ? (resolvePrinterProfileBounds(
            setup,
            printerProfiles,
            selectedProfileId ?? printerProfileId,
          ) ?? null)
        : undefined,
  };
}
