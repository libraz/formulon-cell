/**
 * Shared color palette — the spreadsheet-standard "Office" color picker.
 *
 * A single source of truth for every color surface (ribbon font/fill color,
 * border line color, the Format Cells dialog swatches). The data mirrors the
 * default "Office" document theme: a 10-column theme grid (one base row plus
 * five tint/shade rows) and a 10-color standard row.
 *
 * The component is framework-agnostic vanilla DOM so the core, the React
 * wrapper and the Vue wrapper all consume the exact same widget.
 */

/** One column of the theme grid: a base color and its five tint/shade steps. */
export interface ThemeColorColumn {
  /** Human-readable role, e.g. "Blue, Accent 1". */
  readonly name: string;
  /** Base color hex (the top row of the theme grid). */
  readonly base: string;
  /** Five tint/shade variants, ordered top → bottom as shown in the grid. */
  readonly variants: readonly string[];
}

/**
 * The ten columns of the "Office" theme. Variant values are the canonical
 * Excel tint/shade steps for each base color.
 */
export const THEME_COLOR_COLUMNS: readonly ThemeColorColumn[] = [
  {
    name: 'White, Background 1',
    base: '#ffffff',
    variants: ['#f2f2f2', '#d9d9d9', '#bfbfbf', '#a6a6a6', '#808080'],
  },
  {
    name: 'Black, Text 1',
    base: '#000000',
    variants: ['#808080', '#595959', '#404040', '#262626', '#0d0d0d'],
  },
  {
    name: 'Light Gray, Background 2',
    base: '#e7e6e6',
    variants: ['#d0cece', '#aeaaaa', '#767171', '#3b3838', '#161616'],
  },
  {
    name: 'Dark Blue, Text 2',
    base: '#44546a',
    variants: ['#d6dce5', '#adb9ca', '#8497b0', '#333f4f', '#222a35'],
  },
  {
    name: 'Blue, Accent 1',
    base: '#4472c4',
    variants: ['#d9e2f3', '#b4c7e7', '#8eaadb', '#2f5496', '#1f3864'],
  },
  {
    name: 'Orange, Accent 2',
    base: '#ed7d31',
    variants: ['#fbe5d5', '#f8cbad', '#f4b183', '#c55a11', '#843c0c'],
  },
  {
    name: 'Gray, Accent 3',
    base: '#a5a5a5',
    variants: ['#ededed', '#dbdbdb', '#c9c9c9', '#7b7b7b', '#525252'],
  },
  {
    name: 'Gold, Accent 4',
    base: '#ffc000',
    variants: ['#fff2cc', '#ffe599', '#ffd966', '#bf8f00', '#7f6000'],
  },
  {
    name: 'Blue, Accent 5',
    base: '#5b9bd5',
    variants: ['#ddebf7', '#bdd7ee', '#9dc3e6', '#2e74b5', '#1f4e78'],
  },
  {
    name: 'Green, Accent 6',
    base: '#70ad47',
    variants: ['#e2efda', '#c5e0b3', '#a8d08d', '#548235', '#375623'],
  },
] as const;

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
    const label = name ? `${name} (${normalized})` : normalized;
    swatch.title = label;
    swatch.setAttribute('aria-label', label);
    const index = swatches.length;
    swatch.addEventListener('click', () => {
      selected = normalized;
      syncSelection();
      syncRoving();
      options.onPick(normalized);
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
  // Row 0: base colors.
  for (const column of THEME_COLOR_COLUMNS) {
    const cell = makeSwatch(column.base, column.name);
    cell.classList.add('fc-colorpalette__swatch--base');
    themeGrid.appendChild(cell);
  }
  // Rows 1–5: tint/shade variants, kept column-major in the source data but
  // appended row-major so the flat `swatches` array stays grid-aligned.
  for (let row = 0; row < 5; row += 1) {
    for (const column of THEME_COLOR_COLUMNS) {
      const variant = column.variants[row];
      if (variant) themeGrid.appendChild(makeSwatch(variant, column.name));
    }
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
