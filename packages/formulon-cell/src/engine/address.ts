import type { Addr } from './types.js';

export const addrKey = (a: Addr): string => `${a.sheet}:${a.row}:${a.col}`;
