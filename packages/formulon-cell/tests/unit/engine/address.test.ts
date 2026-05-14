import { describe, expect, it } from 'vitest';

import { addrKey } from '../../../src/engine/address.js';

describe('engine/address', () => {
  it('emits a sheet:row:col string key', () => {
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).toBe('0:0:0');
    expect(addrKey({ sheet: 2, row: 5, col: 8 })).toBe('2:5:8');
  });

  it('distinguishes addresses that differ only in one component', () => {
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).not.toBe(addrKey({ sheet: 1, row: 0, col: 0 }));
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).not.toBe(addrKey({ sheet: 0, row: 1, col: 0 }));
    expect(addrKey({ sheet: 0, row: 0, col: 0 })).not.toBe(addrKey({ sheet: 0, row: 0, col: 1 }));
  });

  it('handles the worksheet upper bound', () => {
    expect(addrKey({ sheet: 0, row: 1048575, col: 16383 })).toBe('0:1048575:16383');
  });
});
