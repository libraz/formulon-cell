/**
 * Shared color palette — the spreadsheet-standard "Office" color picker.
 *
 * A single source of truth for every color surface (ribbon font/fill color,
 * border line color, the Format Cells dialog swatches). The theme grid shows a
 * workbook theme's ten colours, each over five tint/shade rows, and a 10-color
 * standard row follows; the default theme stands in until a theme is given.
 *
 * The component is framework-agnostic vanilla DOM so the core, the React
 * wrapper and the Vue wrapper all consume the exact same widget.
 */

import { argbFromCssColor, cssColorFromArgb, tintCssColor } from '../engine/theme-color.js';

/** One column of the theme grid: a base color and its five tint/shade steps. */
export interface ThemeColorColumn {
  /** The colour's role in the theme, e.g. "Accent 1". */
  readonly name: string;
  /** Base color hex (the top row of the theme grid). */
  readonly base: string;
  /** Five tint/shade variants, ordered top → bottom as shown in the grid. */
  readonly variants: readonly string[];
  /** Each variant's name, e.g. "Accent 1, Lighter 80%". */
  readonly variantNames: readonly string[];
}

/** The variants of a base colour by its luminance, as tints top to bottom. */
const variantTints = (base: string): readonly number[] => {
  const argb = argbFromCssColor(base) ?? 0;
  const channels = [(argb >>> 16) & 0xff, (argb >>> 8) & 0xff, argb & 0xff];
  const lum = (Math.max(...channels) + Math.min(...channels)) / 510;
  if (lum === 0) return [0.5, 0.35, 0.25, 0.15, 0.05];
  if (lum === 1) return [-0.05, -0.15, -0.25, -0.35, -0.5];
  if (lum < 0.2) return [0.9, 0.75, 0.5, 0.25, 0.1];
  if (lum > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9];
  return [0.8, 0.6, 0.4, -0.25, -0.5];
};

const variantName = (tint: number): string =>
  `${tint > 0 ? 'Lighter' : 'Darker'} ${Math.round(Math.abs(tint) * 100)}%`;

/** The grid's columns in clrScheme slots (dk1, lt1, dk2, lt2, accent1..6), with their roles. */
const GRID_SLOTS: readonly (readonly [slot: number, role: string])[] = [
  [1, 'Background 1'],
  [0, 'Text 1'],
  [3, 'Background 2'],
  [2, 'Text 2'],
  [4, 'Accent 1'],
  [5, 'Accent 2'],
  [6, 'Accent 3'],
  [7, 'Accent 4'],
  [8, 'Accent 5'],
  [9, 'Accent 6'],
];

/** The default theme's colours in clrScheme order. */
const DEFAULT_THEME_COLORS: readonly number[] = [
  0xff000000, 0xffffffff, 0xff0e2841, 0xffe8e8e8, 0xff156082, 0xffe97132, 0xff196b24, 0xff0f9ed5,
  0xffa02b93, 0xff4ea72e,
];

/**
 * The theme grid of a workbook theme's ARGB colours in clrScheme order: each
 * background, text and accent colour over its five variants, tinted as the
 * spreadsheet paints them. Falls back to the default theme for a short list.
 */
export function themeColorColumns(colors: readonly number[] | undefined): ThemeColorColumn[] {
  const source = colors && colors.length >= 10 ? colors : DEFAULT_THEME_COLORS;
  return GRID_SLOTS.map(([slot, role]) => {
    const base = cssColorFromArgb(source[slot] ?? 0);
    const tints = variantTints(base);
    return {
      name: role,
      base,
      variants: tints.map((tint) => tintCssColor(base, tint) ?? base),
      variantNames: tints.map((tint) => `${role}, ${variantName(tint)}`),
    };
  });
}

/** The theme grid of the default theme. */
export const THEME_COLOR_COLUMNS: readonly ThemeColorColumn[] = themeColorColumns(undefined);

/** The ten "Standard Colors" shown below the theme grid. */
export const STANDARD_COLORS: readonly string[] = [
  '#c00000',
  '#ff0000',
  '#ffc000',
  '#ffff00',
  '#92d050',
  '#00b050',
  '#00b0f0',
  '#0070c0',
  '#002060',
  '#7030a0',
] as const;

/** Number of columns in both the theme grid and the standard row. */
export const PALETTE_COLUMNS = 10;

/** Normalize a hex string for equality checks (lowercase, leading `#`). */
export function normalizeHex(hex: string): string {
  const trimmed = hex.trim().toLowerCase();
  const value = trimmed.startsWith('#') ? trimmed : `#${trimmed}`;
  // Expand shorthand #abc → #aabbcc so comparisons are stable.
  if (/^#[0-9a-f]{3}$/.test(value)) {
    return `#${value[1]}${value[1]}${value[2]}${value[2]}${value[3]}${value[3]}`;
  }
  return value;
}

/** Optional "Automatic" entry shown above the theme grid. */
export interface AutomaticColorOption {
  /** Button label, e.g. "Automatic". */
  readonly label: string;
  /** Color applied when the button is chosen, e.g. "#000000". */
  readonly color: string;
}

export interface ColorPaletteOptions {
  /** Heading for the theme grid, e.g. "Theme Colors". */
  readonly themeLabel: string;
  /** Heading for the standard row, e.g. "Standard Colors". */
  readonly standardLabel: string;
  /** The theme grid, from `themeColorColumns`; the default theme's when unset. */
  readonly theme?: readonly ThemeColorColumn[];
  /** Currently selected color; the matching swatch is highlighted. */
  readonly value?: string | null;
  /** When set, renders an "Automatic" button above the theme grid. */
  readonly automatic?: AutomaticColorOption | null;
  /** Optional disabled Excel compatibility row shown above fill-color palettes. */
  readonly highContrastOnlyLabel?: string | null;
  /** When set, renders a "More Colors…" trigger below the standard row. */
  readonly moreColorsLabel?: string | null;
  /** Accessible label for the whole palette group. */
  readonly ariaLabel?: string;
  /** Fired when a swatch (or the Automatic button) is chosen. */
  onPick(color: string): void;
  /** Fired when the "More Colors…" trigger is activated. */
  onMoreColors?(): void;
}

export interface ColorPaletteHandle {
  /** Root element to insert into the DOM. */
  readonly el: HTMLElement;
  /** Update the highlighted swatch (e.g. after the selection changes). */
  setValue(color: string | null): void;
  /** Move keyboard focus to the selected swatch, or the first swatch. */
  focus(): void;
}

/** The theme grid's colours and names row-major: the bases, then each variant row. */
const themeGridCells = (
  columns: readonly ThemeColorColumn[],
): (readonly [color: string, name: string])[] => [
  ...columns.map((column) => [column.base, column.name] as const),
  ...[0, 1, 2, 3, 4].flatMap((row) =>
    columns.map(
      (column) =>
        [column.variants[row] ?? column.base, column.variantNames[row] ?? column.name] as const,
    ),
  ),
];

const swatchLabel = (color: string, name: string): string => (name ? `${name} (${color})` : color);

/**
 * Show `columns` in the theme grid of every palette under `root`, for a workbook
 * whose theme changed after the palette was built. A swatch value copied into
 * another `data-` attribute of the swatch moves with it.
 */
export function syncColorPaletteTheme(
  root: ParentNode,
  columns: readonly ThemeColorColumn[],
): void {
  const cells = themeGridCells(columns);
  for (const grid of root.querySelectorAll<HTMLElement>('.fc-colorpalette__grid--theme')) {
    grid.querySelectorAll<HTMLButtonElement>('.fc-colorpalette__swatch').forEach((swatch, i) => {
      const cell = cells[i];
      if (!cell) return;
      const color = normalizeHex(cell[0]);
      const previous = swatch.dataset.color;
      for (const key of Object.keys(swatch.dataset)) {
        if (swatch.dataset[key] === previous) swatch.dataset[key] = color;
      }
      swatch.style.backgroundColor = color;
      swatch.title = swatchLabel(color, cell[1]);
      swatch.setAttribute('aria-label', swatch.title);
    });
  }
}

/** A multi-hue ring — the "More Colors" affordance. */
function createColorWheelIcon(): HTMLSpanElement {
  const wheel = document.createElement('span');
  wheel.className = 'fc-colorpalette__wheel';
  wheel.setAttribute('aria-hidden', 'true');
  return wheel;
}

/**
 * Build the shared color palette widget.
 *
 * Keyboard model: a single roving-tabindex grid spans the theme rows and the
 * standard row (both 10 wide), so arrow keys flow seamlessly between the two.
 */
export function createColorPalette(options: ColorPaletteOptions): ColorPaletteHandle {
  const root = document.createElement('div');
  root.className = 'fc-colorpalette';
  root.setAttribute('role', 'group');
  if (options.ariaLabel) root.setAttribute('aria-label', options.ariaLabel);

  /** Every swatch button, row-major, used for roving-tabindex navigation. */
  const swatches: HTMLButtonElement[] = [];
  let selected = options.value ? normalizeHex(options.value) : null;

  const syncSelection = (): void => {
    for (const swatch of swatches) {
      const isSelected = swatch.dataset.color === selected;
      swatch.setAttribute('aria-pressed', isSelected ? 'true' : 'false');
      swatch.classList.toggle('fc-colorpalette__swatch--selected', isSelected);
    }
  };

  /** Index of the swatch that owns tabindex=0 (the roving focus target). */
  const rovingIndex = (): number => {
    const idx = swatches.findIndex((s) => s.dataset.color === selected);
    return idx >= 0 ? idx : 0;
  };

  const syncRoving = (): void => {
    const active = rovingIndex();
    swatches.forEach((swatch, i) => {
      swatch.tabIndex = i === active ? 0 : -1;
    });
  };

  const focusSwatch = (index: number): void => {
    const clamped = Math.max(0, Math.min(swatches.length - 1, index));
    const target = swatches[clamped];
    if (!target) return;
    for (const swatch of swatches) swatch.tabIndex = -1;
    target.tabIndex = 0;
    target.focus();
  };

  const onSwatchKeydown = (event: KeyboardEvent, index: number): void => {
    let next = index;
    switch (event.key) {
      case 'ArrowRight':
        next = index + 1;
        break;
      case 'ArrowLeft':
        next = index - 1;
        break;
      case 'ArrowDown':
        next = index + PALETTE_COLUMNS;
        break;
      case 'ArrowUp':
        next = index - PALETTE_COLUMNS;
        break;
      case 'Home':
        next = event.ctrlKey ? 0 : index - (index % PALETTE_COLUMNS);
        break;
      case 'End':
        next = event.ctrlKey
          ? swatches.length - 1
          : index - (index % PALETTE_COLUMNS) + (PALETTE_COLUMNS - 1);
        break;
      default:
        return;
    }
    event.preventDefault();
    if (next < 0 || next >= swatches.length) return;
    focusSwatch(next);
  };

  const makeSwatch = (color: string, name: string): HTMLButtonElement => {
    const normalized = normalizeHex(color);
    const swatch = document.createElement('button');
    swatch.type = 'button';
    swatch.className = 'fc-colorpalette__swatch';
    swatch.dataset.color = normalized;
    swatch.style.backgroundColor = normalized;
    swatch.setAttribute('aria-pressed', 'false');
    const label = swatchLabel(normalized, name);
    swatch.title = label;
    swatch.setAttribute('aria-label', label);
    const index = swatches.length;
    // The colour is read at the click: a theme change rewrites the swatch.
    swatch.addEventListener('click', () => {
      const color = swatch.dataset.color ?? normalized;
      selected = color;
      syncSelection();
      syncRoving();
      options.onPick(color);
    });
    swatch.addEventListener('keydown', (event) => onSwatchKeydown(event, index));
    swatches.push(swatch);
    return swatch;
  };

  // ── Automatic ─────────────────────────────────────────────────────────
  if (options.highContrastOnlyLabel) {
    const highContrast = document.createElement('label');
    highContrast.className = 'fc-colorpalette__contrast';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.setAttribute('disabled', '');
    input.tabIndex = -1;
    const text = document.createElement('span');
    text.textContent = options.highContrastOnlyLabel;
    highContrast.append(input, text);
    root.appendChild(highContrast);
  }

  // ── Automatic ─────────────────────────────────────────────────────────
  if (options.automatic) {
    const auto = document.createElement('button');
    auto.type = 'button';
    auto.className = 'fc-colorpalette__action fc-colorpalette__action--automatic';
    const chip = document.createElement('span');
    chip.className = 'fc-colorpalette__action-chip';
    chip.style.backgroundColor = options.automatic.color;
    const text = document.createElement('span');
    text.textContent = options.automatic.label;
    auto.append(chip, text);
    const autoColor = options.automatic.color;
    auto.addEventListener('click', () => options.onPick(autoColor));
    root.appendChild(auto);
  }

  // ── Theme colors ──────────────────────────────────────────────────────
  const themeHeading = document.createElement('div');
  themeHeading.className = 'fc-colorpalette__heading';
  themeHeading.textContent = options.themeLabel;
  root.appendChild(themeHeading);

  const themeGrid = document.createElement('div');
  themeGrid.className = 'fc-colorpalette__grid fc-colorpalette__grid--theme';
  for (const [color, name] of themeGridCells(options.theme ?? THEME_COLOR_COLUMNS)) {
    const cell = makeSwatch(color, name);
    if (themeGrid.childElementCount < PALETTE_COLUMNS) {
      cell.classList.add('fc-colorpalette__swatch--base');
    }
    themeGrid.appendChild(cell);
  }
  root.appendChild(themeGrid);

  // ── Standard colors ───────────────────────────────────────────────────
  const standardHeading = document.createElement('div');
  standardHeading.className = 'fc-colorpalette__heading';
  standardHeading.textContent = options.standardLabel;
  root.appendChild(standardHeading);

  const standardGrid = document.createElement('div');
  standardGrid.className = 'fc-colorpalette__grid fc-colorpalette__grid--standard';
  for (const color of STANDARD_COLORS) {
    standardGrid.appendChild(makeSwatch(color, ''));
  }
  root.appendChild(standardGrid);

  // ── More colors ───────────────────────────────────────────────────────
  if (options.moreColorsLabel) {
    const more = document.createElement('button');
    more.type = 'button';
    more.className = 'fc-colorpalette__action fc-colorpalette__action--more';
    more.append(createColorWheelIcon());
    const text = document.createElement('span');
    text.textContent = options.moreColorsLabel;
    more.appendChild(text);
    more.addEventListener('click', () => options.onMoreColors?.());
    root.appendChild(more);
  }

  syncSelection();
  syncRoving();

  return {
    el: root,
    setValue(color: string | null): void {
      selected = color ? normalizeHex(color) : null;
      syncSelection();
      syncRoving();
    },
    focus(): void {
      swatches[rovingIndex()]?.focus();
    },
  };
}
