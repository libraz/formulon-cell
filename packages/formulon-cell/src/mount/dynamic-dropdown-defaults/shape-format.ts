import { selectedSessionIllustration } from '../../commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeOutline,
  shapeFillColor,
  shapeHasFill,
  shapeOutlineColor,
} from '../../commands/shape-format.js';
import { normalizeHex } from '../../components/color-palette.js';
import type { SpreadsheetInstance } from '../../index.js';
import type { SessionIllustration } from '../../store/store.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled } from './menu-feedback.js';

type ShapeFormatDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyShapeFillAction'
  | 'updateShapeFillMenu'
  | 'applyShapeOutlineAction'
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
    updateShapeOutlineMenu(menu) {
      const item = selectedShape();
      projectColorMenu(menu, 'shapeOutline', item ? shapeOutlineColor(item) : null, !item);
    },
  };
}
