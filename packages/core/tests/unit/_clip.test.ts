import { describe, expect, it } from 'vitest';
describe('happy-dom navigator.clipboard', () => {
  it('exposes clipboard or not?', () => {
    // biome-ignore lint/suspicious/noConsole: probe
    console.log('typeof navigator', typeof navigator);
    // biome-ignore lint/suspicious/noConsole: probe
    console.log('navigator.clipboard', !!(typeof navigator !== 'undefined' && navigator.clipboard));
    // biome-ignore lint/suspicious/noConsole: probe
    console.log('readText fn', typeof navigator?.clipboard?.readText);
    expect(true).toBe(true);
  });
});
