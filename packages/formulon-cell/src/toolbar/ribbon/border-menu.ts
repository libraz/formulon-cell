// Wires the integrated dropdown built by createBordersMenu(): edge / frame
// / combined presets, "More borders..." entry, and the "border-draw" block
// which arms the border-draw controller (drives pointer-edge editing on
// the grid) and exposes two submenus for the line color & line style
// brush settings.

import { type CellBorderStyle, type SpreadsheetInstance, setBorderPreset } from '../../index.js';

import { focusMenuItem, handleMenuKeydown } from '../menu-a11y.js';
import { RIBBON_BORDERS_MENU_ID } from './activation.js';

export interface BorderMenuCtx {
  getInst: () => SpreadsheetInstance | null;
  sheetEl: HTMLElement;
  getSelectedBorderStyle: () => CellBorderStyle;
  setSelectedBorderStyle: (style: CellBorderStyle) => void;
  getSelectedBorderColor: () => string;
  applyRibbonFormat: (
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ) => void;
}

export interface BorderMenuApi {
  openBorderMenu: () => void;
  closeBorderMenu: (restoreFocus?: boolean) => void;
  closeBorderSubmenus: () => void;
  refreshBorderMenuState: () => void;
  applyBorderPresetMenuAction: (key: string) => void;
  applyBorderDrawMenuAction: (action: string | undefined) => void;
  detach: () => void;
}

type BorderPresetKey =
  | 'none'
  | 'outline'
  | 'thickOutline'
  | 'all'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom'
  | 'thickBottom'
  | 'topAndBottom'
  | 'topAndThickBottom'
  | 'topAndDoubleBottom';

// Map menu key -> engine preset. `clear` is the "no border" entry: the
// engine's `'none'` preset wipes every side.
const MENU_TO_PRESET: Record<string, BorderPresetKey> = {
  clear: 'none',
  all: 'all',
  outline: 'outline',
  thickOutline: 'thickOutline',
  top: 'top',
  bottom: 'bottom',
  left: 'left',
  right: 'right',
  doubleBottom: 'doubleBottom',
  thickBottom: 'thickBottom',
  topAndBottom: 'topAndBottom',
  topAndThickBottom: 'topAndThickBottom',
  topAndDoubleBottom: 'topAndDoubleBottom',
};

const BORDER_DRAW_ACTIVE_CLASS = 'fc-tb__menu-item--active';
const BORDER_MENU_ID = RIBBON_BORDERS_MENU_ID;
const BORDER_MENU_SELECTOR = `#${BORDER_MENU_ID}`;

export const createBorderMenu = (ctx: BorderMenuCtx): BorderMenuApi => {
  const borderBtn = document.getElementById('btn-borders');
  const borderMenu = document.getElementById(BORDER_MENU_ID);
  const lineStyleSubmenu =
    borderMenu?.querySelector<HTMLElement>('.fc-tb__submenu--line-style') ?? null;

  const getBorderBtn = (): HTMLButtonElement | null =>
    document.getElementById('btn-borders') as HTMLButtonElement | null;
  const getBorderMenu = (): HTMLDivElement | null =>
    document.getElementById(BORDER_MENU_ID) as HTMLDivElement | null;
  const getLineColorSubmenu = (): HTMLElement | null =>
    getBorderMenu()?.querySelector<HTMLElement>('.fc-tb__submenu--line-color') ?? null;
  const getLineStyleSubmenu = (): HTMLElement | null =>
    getBorderMenu()?.querySelector<HTMLElement>('.fc-tb__submenu--line-style') ?? null;

  const closeBorderSubmenus = (): void => {
    const lineColor = getLineColorSubmenu();
    const lineStyle = getLineStyleSubmenu();
    if (lineColor) lineColor.hidden = true;
    if (lineStyle) lineStyle.hidden = true;
    getBorderMenu()
      ?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]')
      .forEach((b) => {
        b.setAttribute('aria-expanded', 'false');
      });
  };

  const closeBorderMenu = (restoreFocus = false): void => {
    const menu = getBorderMenu();
    const btn = getBorderBtn();
    if (!menu) return;
    menu.hidden = true;
    btn?.setAttribute('aria-expanded', 'false');
    closeBorderSubmenus();
    if (restoreFocus) btn?.focus();
  };

  const refreshBorderMenuState = (): void => {
    const menu = getBorderMenu();
    if (!menu) return;
    // Reflect currently-armed draw mode in the menu so the user can see
    // (and toggle off) the active brush.
    const mode = ctx.getInst()?.borderDraw?.getMode() ?? null;
    menu.querySelectorAll<HTMLButtonElement>('[data-border-draw]').forEach((btn) => {
      const armed = btn.dataset.borderDraw === mode;
      btn.classList.toggle(BORDER_DRAW_ACTIVE_CLASS, armed);
      btn.setAttribute('aria-checked', armed ? 'true' : 'false');
    });
  };

  const openBorderMenu = (): void => {
    const menu = getBorderMenu();
    const btn = getBorderBtn();
    if (!menu) return;
    refreshBorderMenuState();
    menu.hidden = false;
    btn?.setAttribute('aria-expanded', 'true');
    focusMenuItem(menu);
  };

  const onBorderButtonClick = (e: MouseEvent): void => {
    e.stopPropagation();
    if (!borderMenu) return;
    if (borderMenu.hidden) openBorderMenu();
    else closeBorderMenu();
  };
  borderBtn?.addEventListener('click', onBorderButtonClick);

  const onDocumentMouseDown = (e: MouseEvent): void => {
    const menu = getBorderMenu();
    const btn = getBorderBtn();
    if (!menu || menu.hidden) return;
    if (menu.contains(e.target as Node)) return;
    if (btn?.contains(e.target as Node)) return;
    closeBorderMenu();
  };
  document.addEventListener('mousedown', onDocumentMouseDown);

  const onDocumentEscapeKey = (e: KeyboardEvent): void => {
    const menu = getBorderMenu();
    if (e.key === 'Escape' && !menu?.hidden) closeBorderMenu(true);
  };
  document.addEventListener('keydown', onDocumentEscapeKey);

  const onBorderMenuKeydown = (e: KeyboardEvent): void => {
    if (!borderMenu) return;
    handleMenuKeydown(e, borderMenu, { close: closeBorderMenu, restoreFocusTo: borderBtn });
  };
  borderMenu?.addEventListener('keydown', onBorderMenuKeydown);

  const onDocumentSubmenuKeydown = (event: KeyboardEvent): void => {
    const target = event.target as Element | null;
    if (!(target instanceof Element)) return;
    const menu = target?.closest<HTMLDivElement>(BORDER_MENU_SELECTOR);
    if (!menu || menu === borderMenu) return;
    handleMenuKeydown(event, menu, { close: closeBorderMenu, restoreFocusTo: getBorderBtn() });
  };
  document.addEventListener('keydown', onDocumentSubmenuKeydown);

  const applyBorderPresetMenuAction = (key: string): void => {
    const i = ctx.getInst();
    if (!i) return;
    if (key === 'format') {
      closeBorderMenu();
      i.openFormatDialog();
      return;
    }
    const preset = MENU_TO_PRESET[key];
    if (!preset) return;
    closeBorderMenu();
    i.borderDraw?.deactivate();
    ctx.applyRibbonFormat((state, store) =>
      setBorderPreset(state, store, preset, ctx.getSelectedBorderStyle()),
    );
  };

  const applyBorderDrawMenuAction = (action: string | undefined): void => {
    const i = ctx.getInst();
    if (!i) return;
    if (action !== 'draw' && action !== 'grid' && action !== 'erase') return;
    const draw = i.borderDraw;
    if (!draw) return;
    if (draw.getMode() === action) {
      draw.deactivate();
    } else {
      draw.activate(action, ctx.getSelectedBorderStyle(), ctx.getSelectedBorderColor());
    }
    closeBorderMenu();
    refreshBorderMenuState();
    ctx.sheetEl.focus();
  };

  borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-preset]').forEach((btn) => {
    btn.addEventListener('click', () => {
      applyBorderPresetMenuAction(btn.dataset.borderPreset ?? '');
    });
  });

  borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-draw]').forEach((btn) => {
    btn.addEventListener('click', () => {
      applyBorderDrawMenuAction(btn.dataset.borderDraw);
    });
  });

  const openSubmenu = (which: 'lineColor' | 'lineStyle'): void => {
    const menu = getBorderMenu();
    const lineColor = getLineColorSubmenu();
    const lineStyle = getLineStyleSubmenu();
    if (which === 'lineColor') {
      if (lineStyle) lineStyle.hidden = true;
      if (lineColor) lineColor.hidden = false;
    } else {
      if (lineColor) lineColor.hidden = true;
      if (lineStyle) lineStyle.hidden = false;
    }
    menu?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]').forEach((b) => {
      b.setAttribute('aria-expanded', b.dataset.borderSubmenu === which ? 'true' : 'false');
    });
  };

  borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]').forEach((btn) => {
    btn.addEventListener('mouseenter', () => {
      const which = btn.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
      if (which) openSubmenu(which);
    });
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const which = btn.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
      if (which) openSubmenu(which);
    });
  });

  // Mousing over a non-submenu item dismisses any open submenu — matches
  // the single-active-submenu behavior.
  borderMenu
    ?.querySelectorAll<HTMLButtonElement>('[data-border-preset], [data-border-draw]')
    .forEach((btn) => {
      btn.addEventListener('mouseenter', closeBorderSubmenus);
    });

  // Line-color picks are handled by the shared palette's onPick callback,
  // wired in createLineColorSubmenu().

  lineStyleSubmenu
    ?.querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
    .forEach((styleBtn) => {
      styleBtn.addEventListener('click', () => {
        const value = styleBtn.dataset.borderLineStyle ?? 'thin';
        if (value !== 'none') {
          const next = value as CellBorderStyle;
          ctx.setSelectedBorderStyle(next);
          ctx.getInst()?.borderDraw?.setStyle(next);
        }
        lineStyleSubmenu
          .querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
          .forEach((s) => {
            s.setAttribute('aria-checked', s === styleBtn ? 'true' : 'false');
          });
        closeBorderSubmenus();
      });
    });

  const onDocumentClick = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    if (!(target instanceof Element)) return;
    const menu = target?.closest<HTMLElement>(BORDER_MENU_SELECTOR);
    if (!menu || menu === borderMenu) return;
    const preset = target?.closest<HTMLButtonElement>('[data-border-preset]');
    if (preset) {
      event.preventDefault();
      applyBorderPresetMenuAction(preset.dataset.borderPreset ?? '');
      return;
    }
    const draw = target?.closest<HTMLButtonElement>('[data-border-draw]');
    if (draw) {
      event.preventDefault();
      applyBorderDrawMenuAction(draw.dataset.borderDraw);
      return;
    }
    const submenu = target?.closest<HTMLButtonElement>('[data-border-submenu]');
    if (submenu) {
      event.preventDefault();
      const which = submenu.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
      if (which) openSubmenu(which);
      return;
    }
    const lineStyle = target?.closest<HTMLButtonElement>('[data-border-line-style]');
    if (lineStyle) {
      event.preventDefault();
      const value = lineStyle.dataset.borderLineStyle ?? 'thin';
      const lineStyleMenu = getLineStyleSubmenu();
      if (value !== 'none') {
        const next = value as CellBorderStyle;
        ctx.setSelectedBorderStyle(next);
        ctx.getInst()?.borderDraw?.setStyle(next);
      }
      lineStyleMenu
        ?.querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
        .forEach((s) => {
          s.setAttribute('aria-checked', s === lineStyle ? 'true' : 'false');
        });
      closeBorderSubmenus();
    }
  };
  document.addEventListener('click', onDocumentClick);

  const onDocumentMouseOver = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    if (!(target instanceof Element)) return;
    const menu = target?.closest<HTMLElement>(BORDER_MENU_SELECTOR);
    if (!menu || menu === borderMenu) return;
    const submenu = target?.closest<HTMLButtonElement>('[data-border-submenu]');
    if (submenu) {
      const which = submenu.dataset.borderSubmenu as 'lineColor' | 'lineStyle' | undefined;
      if (which) openSubmenu(which);
      return;
    }
    if (target?.closest('[data-border-preset], [data-border-draw]')) closeBorderSubmenus();
  };
  document.addEventListener('mouseover', onDocumentMouseOver);

  const detach = (): void => {
    borderBtn?.removeEventListener('click', onBorderButtonClick);
    borderMenu?.removeEventListener('keydown', onBorderMenuKeydown);
    document.removeEventListener('mousedown', onDocumentMouseDown);
    document.removeEventListener('keydown', onDocumentEscapeKey);
    document.removeEventListener('keydown', onDocumentSubmenuKeydown);
    document.removeEventListener('click', onDocumentClick);
    document.removeEventListener('mouseover', onDocumentMouseOver);
  };

  return {
    openBorderMenu,
    closeBorderMenu,
    closeBorderSubmenus,
    refreshBorderMenuState,
    applyBorderPresetMenuAction,
    applyBorderDrawMenuAction,
    detach,
  };
};
