const baseDigits = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
const engineeringBaseValue = (text: string, base: 2 | 8 | 16): number | null => {
  const normalized = text.trim().toUpperCase();
  if (normalized === '' || normalized.length > 10) return null;
  let unsigned = 0;
  for (const char of normalized) {
    const digit = baseDigits.indexOf(char);
    if (digit < 0 || digit >= base) return null;
    unsigned = unsigned * base + digit;
  }
  const signThreshold = base ** 9;
  const modulus = base ** 10;
  return normalized.length === 10 && unsigned >= signThreshold ? unsigned - modulus : unsigned;
};
const engineeringBaseText = (
  value: number,
  base: 2 | 8 | 16,
  places: number | null,
): string | null => {
  const negativeLimit = -(base ** 9);
  const positiveLimit = base ** 9 - 1;
  if (value < negativeLimit || value > positiveLimit || (places !== null && places < 0)) {
    return null;
  }
  if (value < 0)
    return Math.trunc(value + base ** 10)
      .toString(base)
      .toUpperCase();
  const text = Math.trunc(value).toString(base).toUpperCase();
  if (places !== null && text.length > places) return null;
  return places === null ? text : text.padStart(places, '0');
};
const romanNumerals: [number, string][] = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];
const romanText = (value: number): string => {
  let remaining = value;
  let result = '';
  for (const [amount, symbol] of romanNumerals) {
    while (remaining >= amount) {
      result += symbol;
      remaining -= amount;
    }
  }
  return result;
};
const romanValue = (value: string): number | null => {
  const normalized = value.trim().toUpperCase();
  if (normalized === '') return null;
  let index = 0;
  let result = 0;
  while (index < normalized.length) {
    const match = romanNumerals.find(([, symbol]) => normalized.startsWith(symbol, index));
    if (!match) return null;
    result += match[0];
    index += match[1].length;
  }
  return romanText(result) === normalized ? result : null;
};
const maxBitValue = 281_474_976_710_655;
const bitOperand = (value: number): bigint | null => {
  const integer = Math.trunc(value);
  return integer < 0 || integer > maxBitValue ? null : BigInt(integer);
};

export {
  baseDigits,
  bitOperand,
  engineeringBaseText,
  engineeringBaseValue,
  maxBitValue,
  romanText,
  romanValue,
};
