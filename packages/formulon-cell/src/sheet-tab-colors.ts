import type { Strings } from './i18n/strings.js';

export type SheetTabColorChoice = {
  key: string;
  action: string;
  wrapperAction:
    | 'tabColorNone'
    | 'tabColorRed'
    | 'tabColorOrange'
    | 'tabColorYellow'
    | 'tabColorGreen'
    | 'tabColorBlue'
    | 'tabColorPurple'
    | 'tabColorGray';
  color: string | null;
  labelKey:
    | 'noColor'
    | 'tabColorRed'
    | 'tabColorOrange'
    | 'tabColorYellow'
    | 'tabColorGreen'
    | 'tabColorBlue'
    | 'tabColorPurple'
    | 'tabColorGray';
};

export const SHEET_TAB_COLOR_CHOICES = [
  {
    key: 'none',
    action: 'tab-color-none',
    wrapperAction: 'tabColorNone',
    color: null,
    labelKey: 'noColor',
  },
  {
    key: 'red',
    action: 'tab-color-red',
    wrapperAction: 'tabColorRed',
    color: '#c00000',
    labelKey: 'tabColorRed',
  },
  {
    key: 'orange',
    action: 'tab-color-orange',
    wrapperAction: 'tabColorOrange',
    color: '#ed7d31',
    labelKey: 'tabColorOrange',
  },
  {
    key: 'yellow',
    action: 'tab-color-yellow',
    wrapperAction: 'tabColorYellow',
    color: '#ffc000',
    labelKey: 'tabColorYellow',
  },
  {
    key: 'green',
    action: 'tab-color-green',
    wrapperAction: 'tabColorGreen',
    color: '#70ad47',
    labelKey: 'tabColorGreen',
  },
  {
    key: 'blue',
    action: 'tab-color-blue',
    wrapperAction: 'tabColorBlue',
    color: '#4472c4',
    labelKey: 'tabColorBlue',
  },
  {
    key: 'purple',
    action: 'tab-color-purple',
    wrapperAction: 'tabColorPurple',
    color: '#7030a0',
    labelKey: 'tabColorPurple',
  },
  {
    key: 'gray',
    action: 'tab-color-gray',
    wrapperAction: 'tabColorGray',
    color: '#7f7f7f',
    labelKey: 'tabColorGray',
  },
] as const satisfies readonly SheetTabColorChoice[];

export const sheetTabColorByAction = (action: string): string | null | undefined =>
  SHEET_TAB_COLOR_CHOICES.find((choice) => choice.action === action)?.color;

export const sheetTabColorActionForColor = (color: string | undefined): string => {
  const normalized = color?.trim().toLowerCase();
  if (!normalized) return 'tab-color-none';
  return (
    SHEET_TAB_COLOR_CHOICES.find((choice) => choice.color?.toLowerCase() === normalized)?.action ??
    'tab-color-none'
  );
};

export const sheetTabColorChoiceLabel = (
  choice: SheetTabColorChoice,
  strings: Strings['sheetTabs'],
): string => strings[choice.labelKey];
