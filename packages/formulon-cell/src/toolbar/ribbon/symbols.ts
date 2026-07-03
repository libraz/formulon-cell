import type { ToolbarMenuText } from '../../index.js';

const SYMBOL_GROUP_DEFS = [
  {
    labelKey: 'symbolMath',
    symbols: ['±', '×', '÷', '≤', '≥', '≠', '≈', '∞', '√', '∑', '∫', 'π'],
  },
  {
    labelKey: 'symbolGreek',
    symbols: ['Α', 'Β', 'Γ', 'Δ', 'Θ', 'Λ', 'Ξ', 'Π', 'Σ', 'Φ', 'Ψ', 'Ω'],
  },
  { labelKey: 'symbolCurrency', symbols: ['$', '€', '¥', '£', '¢', '₩', '₹', '₽'] },
  { labelKey: 'symbolLegal', symbols: ['©', '®', '™', '§', '¶', '†', '‡', '•'] },
] as const satisfies readonly {
  labelKey: keyof Pick<
    ToolbarMenuText,
    'symbolMath' | 'symbolGreek' | 'symbolCurrency' | 'symbolLegal'
  >;
  symbols: readonly string[];
}[];

export type ToolbarSymbolGroup = {
  label: string;
  symbols: readonly string[];
};

export const toolbarSymbolGroups = (text: ToolbarMenuText): ToolbarSymbolGroup[] =>
  SYMBOL_GROUP_DEFS.map((group) => ({
    label: text[group.labelKey],
    symbols: group.symbols,
  }));
