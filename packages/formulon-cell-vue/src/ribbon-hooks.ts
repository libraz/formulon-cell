import type { RibbonHooks } from '@libraz/formulon-cell/toolbar';

type HookGroups = Record<string, Record<string, unknown> | undefined>;

/** The hooks present in `hooks`, as `group.name` keys; the ribbon remounts only when it changes. */
export const ribbonHooksShape = (hooks: RibbonHooks | undefined): string =>
  Object.entries((hooks ?? {}) as HookGroups)
    .flatMap(([group, members]) =>
      Object.entries(members ?? {})
        .filter(([, fn]) => typeof fn === 'function')
        .map(([name]) => `${group}.${name}`),
    )
    .sort()
    .join(',');

/** Hooks with the shape of `get()` that call whatever `get()` holds at click time. */
export const forwardRibbonHooks = (get: () => RibbonHooks | undefined): RibbonHooks => {
  const out: Record<string, Record<string, unknown>> = {};
  for (const key of ribbonHooksShape(get()).split(',')) {
    if (!key) continue;
    const [group, name] = key.split('.') as [string, string];
    out[group] ??= {};
    out[group][name] = (...args: unknown[]) => {
      const fn = (get() as HookGroups | undefined)?.[group]?.[name];
      return typeof fn === 'function' ? fn(...args) : undefined;
    };
  }
  return out as RibbonHooks;
};
