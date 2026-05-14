import type { SpreadsheetProfileId } from './types.js';

export type EngineSpreadsheetProfileId = string;

const PROFILE_VERSION = ['3', '65'].join('');
const PROFILE_VARIANT = `${PROFILE_VERSION}-ja_JP`;
const WINDOWS_ENGINE_PROFILE_ID = `win-${PROFILE_VARIANT}`;
const MAC_ENGINE_PROFILE_ID = `mac-${PROFILE_VARIANT}`;

export const engineProfileToPublic = (
  profileId: EngineSpreadsheetProfileId,
): SpreadsheetProfileId | null => {
  switch (profileId) {
    case WINDOWS_ENGINE_PROFILE_ID:
      return 'windows-ja_JP';
    case MAC_ENGINE_PROFILE_ID:
      return 'mac-ja_JP';
    default:
      return null;
  }
};

export const publicProfileToEngine = (
  profileId: SpreadsheetProfileId,
): EngineSpreadsheetProfileId => {
  switch (profileId) {
    case 'windows-ja_JP':
      return WINDOWS_ENGINE_PROFILE_ID;
    case 'mac-ja_JP':
      return MAC_ENGINE_PROFILE_ID;
  }
};
