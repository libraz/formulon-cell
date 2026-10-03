import type { CellValue } from '../../engine/types.js';

const textValue = (value: CellValue): string | null => {
  if (value.kind === 'text') return value.value;
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  if (value.kind === 'blank') return '';
  return null;
};
const booleanValue = (value: CellValue): boolean | null => {
  if (value.kind === 'bool') return value.value;
  if (value.kind === 'number' && Number.isFinite(value.value)) return value.value !== 0;
  return null;
};
const nonNegativeInteger = (value: CellValue): number | null => {
  if (value.kind !== 'number' || !Number.isFinite(value.value) || value.value < 0) return null;
  return Math.floor(value.value);
};
const positiveInteger = (value: CellValue): number | null => {
  if (value.kind !== 'number' || !Number.isFinite(value.value) || value.value < 1) return null;
  return Math.floor(value.value);
};
const readNumber = (value: CellValue): number | null =>
  value.kind === 'number' && Number.isFinite(value.value) ? value.value : null;
const readLogical = (value: CellValue): boolean | null => {
  if (value.kind === 'bool') return value.value;
  if (value.kind === 'number' && Number.isFinite(value.value)) return value.value !== 0;
  return null;
};

export { booleanValue, nonNegativeInteger, positiveInteger, readLogical, readNumber, textValue };
