import { selectedSessionIllustration } from '../../commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeOutline,
  applySessionShapeOutlineDash,
  applySessionShapeOutlineWeight,
  SHAPE_LINE_DASHES,
  shapeFillColor,
  shapeHasFill,
  shapeOutlineColor,
  shapeOutlineDash,
  shapeOutlineWeight,
} from '../../commands/shape-format.js';
import { normalizeHex } from '../../components/color-palette.js';
import type { SpreadsheetInstance } from '../../index.js';
import type { SessionIllustration, SessionShapeLineDash } from '../../store/store.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled } from './menu-feedback.js';

type ShapeFormatDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyShapeFillAction'
  | 'updateShapeFillMenu'
  | 'applyShapeOutlineAction'
  | 'applyShapeOutlineWeightAction'
  | 'applyShapeOutlineDashAction'
  | 'updateShapeOutlineMenu'
>;

/** A menu value: a colour, or `none` for No Fill / No Outline. */
const colorOf = (value: string): string | null => (value === 'none' ? null : normalizeHex(value));

/** Check the entry of `current`, a colour or null for the none entry. */
const projectColorMenu = (
  menu: HTMLElement,
  attr: 'shapeFill' | 'shapeOutline',
  current: string | null,
  disabled: boolean,
): void => {
  const selected = current === null ? 'none' : normalizeHex(current);
  for (const button of menu.querySelectorAll<HTMLButtonElement>(
    `[data-${attr === 'shapeFill' ? 'shape-fill' : 'shape-outline'}]`,
  )) {
    const value = button.dataset[attr] ?? '';
    const checked = !disabled && value === selected;
    setMenuControlDisabled(button, disabled);
    if (value === 'none') {
      button.setAttribute('aria-checked', String(checked));
      button.classList.toggle('fc-tb__menu-item--active', checked);
    } else {
      button.setAttribute('aria-pressed', String(checked));
      button.classList.toggle('fc-colorpalette__swatch--selected', checked);
    }
  }
};

/** Check the Weight or Dashes entry whose value is `current`; none while the outline is off. */
const projectRadioPanel = (
  menu: HTMLElement,
  attr: 'shapeOutlineWeight' | 'shapeOutlineDash',
  current: string | null,
  disabled: boolean,
): void => {
  const selector =
    attr === 'shapeOutlineWeight' ? '[data-shape-outline-weight]' : '[data-shape-outline-dash]';
  for (const button of menu.querySelectorAll<HTMLButtonElement>(selector)) {
    const checked = !disabled && button.dataset[attr] === current;
    setMenuControlDisabled(button, disabled);
    button.setAttribute('aria-checked', String(checked));
    button.classList.toggle('fc-tb__menu-item--active', checked);
  }
  for (const trigger of menu.querySelectorAll<HTMLButtonElement>('[data-format-submenu]')) {
    setMenuControlDisabled(trigger, disabled);
  }
};

const isLineDash = (value: string): value is SessionShapeLineDash =>
  (SHAPE_LINE_DASHES as readonly string[]).includes(value);

export function createShapeFormatDropdownDefaults(
  instance: SpreadsheetInstance,
): ShapeFormatDropdownDefaults {
  const selectedShape = (): SessionIllustration | null => {
    const item = selectedSessionIllustration(instance.store.getState());
    return item?.kind === 'shape' ? item : null;
  };

  return {
    applyShapeFillAction(value) {
      const item = selectedShape();
      if (!item) return;
      applySessionShapeFill(
        instance.store,
        item,
        colorOf(value),
        instance.history,
        instance.workbook,
      );
    },
    updateShapeFillMenu(menu) {
      const item = selectedShape();
      const fillable = item !== null && shapeHasFill(item);
      projectColorMenu(menu, 'shapeFill', item ? shapeFillColor(item) : null, !fillable);
    },
    applyShapeOutlineAction(value) {
      const item = selectedShape();
      if (!item) return;
      applySessionShapeOutline(
        instance.store,
        item,
        colorOf(value),
        instance.history,
        instance.workbook,
      );
    },
    applyShapeOutlineWeightAction(value) {
      const item = selectedShape();
      if (!item) return;
      applySessionShapeOutlineWeight(
        instance.store,
        item,
        Number(value),
        instance.history,
        instance.workbook,
      );
    },
    applyShapeOutlineDashAction(value) {
      const item = selectedShape();
      if (!item || !isLineDash(value)) return;
      applySessionShapeOutlineDash(
        instance.store,
        item,
        value,
        instance.history,
        instance.workbook,
      );
    },
    updateShapeOutlineMenu(menu) {
      const item = selectedShape();
      projectColorMenu(menu, 'shapeOutline', item ? shapeOutlineColor(item) : null, !item);
      const weight = item ? shapeOutlineWeight(item) : null;
      projectRadioPanel(menu, 'shapeOutlineWeight', weight === null ? null : String(weight), !item);
      const dash = item && weight !== null ? shapeOutlineDash(item) : null;
      projectRadioPanel(menu, 'shapeOutlineDash', dash, !item);
    },
  };
}
