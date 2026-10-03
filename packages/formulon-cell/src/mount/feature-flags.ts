import { ALL_FEATURE_IDS } from '../extensions/features.js';
import { resolveFlags } from '../extensions/index.js';
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

/** Resolve requested features into final flags, applying policy and context-menu overrides. */
export function resolveMountFlags(
  requested: Parameters<typeof resolveFlags>[0],
  restricted: boolean,
  contextMenuOptions: MountOptions['contextMenu'],
): MountFlags {
  const flags = restrictedFlags(resolveFlags(requested), restricted);
  if (contextMenuOptions && contextMenuOptions.mode !== 'disabled') flags.contextMenu = true;
  if (contextMenuOptions?.mode === 'disabled') flags.contextMenu = false;
  return flags;
}
