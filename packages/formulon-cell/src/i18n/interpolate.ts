/** Replace each `{name}` in `template` with `vars.name`; a missing name becomes the empty string. */
export const interpolate = (template: string, vars: Record<string, string | number>): string =>
  template.replace(/\{(\w+)\}/g, (_, key: string) => String(vars[key] ?? ''));
