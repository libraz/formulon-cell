import { ALL_FEATURE_IDS, type FeatureId, resolveFlags } from '../extensions/features.js';
import type { ResolvedSpreadsheetPlatform } from '../extensions/ui-options.js';
import type { MountOptions } from './types.js';

type MountFlags = ReturnType<typeof resolveFlags>;

// These are the routes integrated with the cell command controller. Other
// mutation surfaces remain unavailable while an interaction policy is active.
export const RESTRICTED_FEATURES = new Set([
  'formulaBar',
  'clipboard',
  'shortcuts',
  'wheel',
  'contextMenu',
]);

export function restrictedFlags(flags: MountFlags, restricted: boolean): MountFlags {
  if (!restricted) return flags;
  return Object.fromEntries(
    ALL_FEATURE_IDS.map((id) => [id, RESTRICTED_FEATURES.has(id) && flags[id]]),
  ) as MountFlags;
}

/** Features a platform's desktop app does not have at all; that profile never
 *  offers them, whatever the host requests. The Mac app has no Quick Analysis. */
const PLATFORM_ABSENT_FEATURES: Record<ResolvedSpreadsheetPlatform, readonly FeatureId[]> = {
  mac: ['quickAnalysis'],
  default: [],
};

/** Whether `platform`'s desktop app lacks `id` altogether. */
export const platformLacksFeature = (
  id: FeatureId,
  platform: ResolvedSpreadsheetPlatform,
): boolean => PLATFORM_ABSENT_FEATURES[platform].includes(id);

/** Resolve requested features into final flags, applying policy, platform and
 *  context-menu overrides. */
export function resolveMountFlags(
  requested: Parameters<typeof resolveFlags>[0],
  restricted: boolean,
  contextMenuOptions: MountOptions['contextMenu'],
  platform: ResolvedSpreadsheetPlatform,
): MountFlags {
  const flags = restrictedFlags(resolveFlags(requested), restricted);
  for (const id of ALL_FEATURE_IDS) if (platformLacksFeature(id, platform)) flags[id] = false;
  if (contextMenuOptions && contextMenuOptions.mode !== 'disabled') flags.contextMenu = true;
  if (contextMenuOptions?.mode === 'disabled') flags.contextMenu = false;
  return flags;
}
