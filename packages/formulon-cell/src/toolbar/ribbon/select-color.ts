// Ribbon `<select>`-style dropdown and color-picker control builders extracted
// from main.ts. The factory wires in i18n, host hooks (apply/read control
// value, font availability, color palette) and produces the DOM elements that
// renderRibbon mounts.

import { createColorPalette } from '../../components/color-palette.js';
import type { SpreadsheetInstance } from '../../mount/types.js';
import type { RibbonCommand } from '../ribbon-model.js';
import {
  FONT_SUBMENU_FAMILIES,
  RECENT_FONT_VALUES,
  THEME_FONT_VALUES,
} from './font-availability.js';
import { createRibbonButton } from './button.js';

export interface SelectColorRibbonText {
  numberFormatNoSpecific: string;
  themeColors: string;
  standardColors: string;
  moreColors: string;
  automatic: string;
  marginsCustomDialog: string;
  marginTop: string;
  marginBottom: string;
  marginLeft: string;
  marginRight: string;
  fontSectionTheme: string;
  fontSectionRecent: string;
  fontSectionAll: string;
  fontRoleHeading: string;
  fontRoleBody: string;
  currentView: string;
}

export interface SelectColorPageScaleText {
  automatic: string;
  page: string;
  pages: string;
}

export interface SelectColorCtx {
  ribbonLang: 'ja' | 'en';
  ribbonText: SelectColorRibbonText;
  pageScaleText: SelectColorPageScaleText;
  getInst: () => SpreadsheetInstance | null;
  applyRibbonControl: (id: string, value: string) => void;
  currentRibbonControlValue: (id: string) => string;
  shouldShowFontOption: (value: string, current: string, locale: 'ja' | 'en') => boolean;
  createRibbonIcon: (name: string) => SVGSVGElement | null;
}

export interface SelectColorApi {
  makeSvg: (viewBox: string, pathData: string, className: string) => SVGSVGElement;
  createRibbonSelect: (command: RibbonCommand) => HTMLDivElement;
  createRibbonColor: (command: RibbonCommand) => HTMLDivElement;
  closeOpenRibbonDropdowns: (except?: HTMLElement) => void;
  updateRibbonSelectDisplay: (wrap: HTMLElement, command: RibbonCommand) => void;
  ribbonSelectLabel: (wrap: HTMLElement, current: string) => string;
  RIBBON_CHEVRON_PATH: string;
}

export const createSelectColorRibbon = (ctx: SelectColorCtx): SelectColorApi => {
  const {
    ribbonLang,
    ribbonText,
    pageScaleText,
    getInst,
    applyRibbonControl,
    currentRibbonControlValue,
    shouldShowFontOption,
    createRibbonIcon,
  } = ctx;

  const makeSvg = (viewBox: string, pathData: string, className: string): SVGSVGElement => {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.classList.add(className);
    svg.setAttribute('viewBox', viewBox);
    svg.setAttribute('fill', 'currentColor');
    svg.setAttribute('focusable', 'false');
    svg.setAttribute('aria-hidden', 'true');
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', pathData);
    svg.appendChild(path);
    return svg;
  };

  const ribbonMarginDetail = (value: string): string | null => {
    const fmt = (top: string, bottom: string, left: string, right: string): string =>
      `${ribbonText.marginTop}: ${top}", ${ribbonText.marginBottom}: ${bottom}", ${ribbonText.marginLeft}: ${left}", ${ribbonText.marginRight}: ${right}"`;
    switch (value) {
      case 'normal':
        return fmt('0.75', '0.75', '0.7', '0.7');
      case 'wide':
        return fmt('1', '1', '1', '1');
      case 'narrow':
        return fmt('0.75', '0.75', '0.25', '0.25');
      case 'custom':
        return ribbonText.marginsCustomDialog;
      default:
        return null;
    }
  };

  const createMarginPresetIcon = (value: string): HTMLSpanElement => {
    const icon = document.createElement('span');
    icon.className = `demo__rb-dd__margin-icon demo__rb-dd__margin-icon--${value}`;
    icon.setAttribute('aria-hidden', 'true');
    icon.append(document.createElement('span'), document.createElement('span'));
    return icon;
  };

  const numberFormatHasSubtitle = (value: string): boolean => value === 'general';

  const numberFormatSubtitle = (value: string): string =>
    value === 'general' ? ribbonText.numberFormatNoSpecific : '';

  const ribbonFontSection = (
    value: string,
    options: readonly { value: string; label: string }[],
  ): string | null => {
    const firstTheme = options.find((option) => THEME_FONT_VALUES.has(option.value))?.value;
    if (value === firstTheme) return ribbonText.fontSectionTheme;
    const firstRecent = options.find((option) => RECENT_FONT_VALUES.has(option.value))?.value;
    if (value === firstRecent) return ribbonText.fontSectionRecent;
    const firstAll = options.find(
      (option) => !THEME_FONT_VALUES.has(option.value) && !RECENT_FONT_VALUES.has(option.value),
    )?.value;
    if (value === firstAll) return ribbonText.fontSectionAll;
    return null;
  };

  const ribbonFontRole = (value: string): string | null => {
    switch (value) {
      case 'Aptos Display':
      case '游ゴシック Light':
        return ribbonText.fontRoleHeading;
      case 'Aptos Narrow':
      case '游ゴシック Regular':
        return ribbonText.fontRoleBody;
      default:
        return null;
    }
  };

  const ribbonOptionsForCommand = (
    command: RibbonCommand,
    current: string,
  ): readonly { value: string; label: string }[] => {
    const options = command.options ?? [];
    if (command.id === 'sheetViewSelect') {
      const inst = getInst();
      const views =
        inst?.store
          .getState()
          .sheetViews.views.filter((view) => view.sheet === inst?.store.getState().data.sheetIndex)
          .map((view) => ({ value: view.id, label: view.name })) ?? [];
      return [{ value: 'current', label: ribbonText.currentView }, ...views];
    }
    if (command.id !== 'fontFamily') return options;
    return options.filter((option) => shouldShowFontOption(option.value, current, ribbonLang));
  };

  const RIBBON_CHEVRON_PATH =
    'M2.15 4.65a.5.5 0 0 1 .7 0L6 7.79l3.15-3.14a.5.5 0 1 1 .7.7l-3.5 3.5a.5.5 0 0 1-.7 0l-3.5-3.5a.5.5 0 0 1 0-.7Z';

  const closeOpenRibbonDropdowns = (except?: HTMLElement): void => {
    for (const open of document.querySelectorAll<HTMLElement>('.demo__rb-dd--open')) {
      if (except && open === except) continue;
      open.classList.remove('demo__rb-dd--open');
      open
        .querySelector<HTMLButtonElement>('.demo__rb-dd__btn')
        ?.setAttribute('aria-expanded', 'false');
      open.querySelector('.demo__rb-dd__list')?.remove();
    }
    for (const open of document.querySelectorAll<HTMLElement>('.demo__rb-color--open')) {
      if (except && open === except) continue;
      open.classList.remove('demo__rb-color--open');
      open
        .querySelector<HTMLButtonElement>('.demo__rb-color__btn')
        ?.setAttribute('aria-expanded', 'false');
      open.querySelector('.demo__color-flyout')?.remove();
    }
  };

  const updateRibbonSelectDisplay = (wrap: HTMLElement, command: RibbonCommand): void => {
    const current = currentRibbonControlValue(command.id);
    const option = ribbonOptionsForCommand(command, current).find(
      (candidate) => candidate.value === current,
    );
    const value = wrap.querySelector<HTMLElement>('.demo__rb-dd__value');
    if (value) {
      const base = option?.label ?? current;
      const role = command.id === 'fontFamily' ? ribbonFontRole(current) : null;
      value.textContent = role ? `${base} ${role}` : base;
    }
  };

  const createRibbonControlButton = (opts: {
    className: string;
    title?: string;
    ariaLabel?: string;
    ariaHaspopup?: string;
    ariaExpanded?: boolean;
    role?: string;
    selected?: boolean;
    tabIndex?: number;
    dataset?: Record<string, string>;
  }): HTMLButtonElement => {
    return createRibbonButton({
      className: opts.className,
      title: opts.title,
      ariaLabel: opts.ariaLabel,
      ariaHaspopup: opts.ariaHaspopup,
      ariaExpanded: opts.ariaExpanded,
      role: opts.role,
      ariaSelected: opts.selected,
      tabIndex: opts.tabIndex,
      dataset: opts.dataset,
    });
  };

  const createRibbonSelect = (command: RibbonCommand): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = `demo__rb-dd${command.className ? ` ${command.className}` : ''}`;
    wrap.dataset.ribbonCommand = command.id;
    wrap.dataset.ribbonSelect = command.id;
    wrap.dataset.ribbonOptions = JSON.stringify(command.options ?? []);

    const button = createRibbonControlButton({
      className: 'demo__rb-dd__btn',
      title: command.title,
      ariaLabel: command.title,
      ariaHaspopup: 'listbox',
      ariaExpanded: false,
    });

    const value = document.createElement('span');
    value.className = 'demo__rb-dd__value';
    button.append(
      value,
      makeSvg(
        '0 0 12 12',
        'M2.15 4.65a.5.5 0 0 1 .7 0L6 7.79l3.15-3.14a.5.5 0 1 1 .7.7l-3.5 3.5a.5.5 0 0 1-.7 0l-3.5-3.5a.5.5 0 0 1 0-.7Z',
        'demo__rb-dd__chev',
      ),
    );
    wrap.appendChild(button);

    let detachDocDown: (() => void) | null = null;
    const close = (): void => {
      wrap.classList.remove('demo__rb-dd--open');
      button.setAttribute('aria-expanded', 'false');
      wrap.querySelector('.demo__rb-dd__list')?.remove();
      detachDocDown?.();
      detachDocDown = null;
    };
    const focusListOption = (list: HTMLElement, index: number): void => {
      const options = Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]'));
      if (options.length === 0) return;
      const next = ((index % options.length) + options.length) % options.length;
      for (const [idx, option] of options.entries()) option.tabIndex = idx === next ? 0 : -1;
      options[next]?.focus({ preventScroll: true });
      options[next]?.scrollIntoView({ block: 'nearest' });
    };
    const pickOption = (option: HTMLButtonElement): void => {
      const nextValue = option.dataset.value;
      if (nextValue == null) return;
      applyRibbonControl(command.id, nextValue);
      const label = option.querySelector<HTMLElement>('.demo__rb-dd__label')?.textContent;
      if (label) value.textContent = label;
      close();
      button.focus({ preventScroll: true });
    };
    const open = (): void => {
      closeOpenRibbonDropdowns(wrap);
      wrap.classList.add('demo__rb-dd--open');
      button.setAttribute('aria-expanded', 'true');
      const list = document.createElement('div');
      list.className = 'demo__rb-dd__list';
      list.setAttribute('role', 'listbox');
      list.setAttribute('aria-label', command.title);
      list.tabIndex = -1;
      const anchorRect = button.getBoundingClientRect();
      list.style.left = `${Math.round(anchorRect.left)}px`;
      list.style.top = `${Math.round(anchorRect.bottom + 3)}px`;
      list.style.minWidth = `${Math.round(anchorRect.width)}px`;
      const current = currentRibbonControlValue(command.id);
      const options = ribbonOptionsForCommand(command, current);
      for (const option of options) {
        const section =
          command.id === 'fontFamily' ? ribbonFontSection(option.value, options) : null;
        if (section) {
          const heading = document.createElement('div');
          heading.className = 'demo__rb-dd__section';
          heading.setAttribute('role', 'presentation');
          heading.textContent = section;
          list.appendChild(heading);
        }
        const selected = option.value === current;
        const item = createRibbonControlButton({
          className: `demo__rb-dd__opt${selected ? ' demo__rb-dd__opt--selected' : ''}`,
          role: 'option',
          selected,
          tabIndex: -1,
          dataset: { value: option.value, fcValue: option.value },
        });
        const check = document.createElement('span');
        check.className = 'demo__rb-dd__check';
        check.setAttribute('aria-hidden', 'true');
        if (selected) {
          check.appendChild(
            makeSvg(
              '0 0 16 16',
              'M13.36 3.74c.29.28.29.77 0 1.05l-7.01 7.01a.75.75 0 0 1-1.06 0L2.64 9.15a.75.75 0 1 1 1.06-1.06l2.12 2.12 6.48-6.47a.75.75 0 0 1 1.06 0Z',
              'demo__rb-dd__check-icon',
            ),
          );
        }
        const label = document.createElement('span');
        label.className = 'demo__rb-dd__label';
        label.textContent = option.label;
        if (command.id === 'marginsPreset') {
          const text = document.createElement('span');
          text.className = 'demo__rb-dd__margin-text';
          const detail = document.createElement('span');
          detail.className = 'demo__rb-dd__detail';
          detail.textContent = ribbonMarginDetail(option.value) ?? '';
          text.append(label, detail);
          item.append(check, createMarginPresetIcon(option.value), text);
        } else if (command.id === 'fontFamily') {
          const preview = document.createElement('span');
          preview.className = 'demo__rb-dd__font-preview';
          preview.style.fontFamily = `"${option.value}", sans-serif`;
          const role = ribbonFontRole(option.value);
          if (role) {
            const detail = document.createElement('span');
            detail.className = 'demo__rb-dd__font-role';
            detail.textContent = role;
            preview.append(label, detail);
          } else {
            preview.append(label);
          }
          item.append(check, preview);
          if (FONT_SUBMENU_FAMILIES.has(option.value)) {
            const arrow = document.createElement('span');
            arrow.className = 'demo__rb-dd__submenu';
            arrow.setAttribute('aria-hidden', 'true');
            arrow.textContent = '›';
            item.appendChild(arrow);
          }
        } else if (command.id === 'numberFormat' && numberFormatHasSubtitle(option.value)) {
          const text = document.createElement('span');
          text.className = 'demo__rb-dd__numfmt-text';
          const detail = document.createElement('span');
          detail.className = 'demo__rb-dd__numfmt-subtitle';
          detail.textContent = numberFormatSubtitle(option.value);
          text.append(label, detail);
          item.append(check, text);
        } else {
          item.append(check, label);
        }
        item.addEventListener('click', () => pickOption(item));
        list.appendChild(item);
      }
      list.addEventListener('keydown', (event) => {
        const options = Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]'));
        const currentIndex = Math.max(
          0,
          options.indexOf(document.activeElement as HTMLButtonElement),
        );
        if (event.key === 'ArrowDown') {
          event.preventDefault();
          focusListOption(list, currentIndex + 1);
        } else if (event.key === 'ArrowUp') {
          event.preventDefault();
          focusListOption(list, currentIndex - 1);
        } else if (event.key === 'Home') {
          event.preventDefault();
          focusListOption(list, 0);
        } else if (event.key === 'End') {
          event.preventDefault();
          focusListOption(list, options.length - 1);
        } else if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          const option = document.activeElement?.closest<HTMLButtonElement>('[role="option"]');
          if (option && list.contains(option)) pickOption(option);
        } else if (event.key === 'Escape') {
          event.preventDefault();
          close();
          button.focus({ preventScroll: true });
        }
      });
      wrap.appendChild(list);
      const selectedIndex = Math.max(
        0,
        Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]')).findIndex(
          (option) => option.getAttribute('aria-selected') === 'true',
        ),
      );
      focusListOption(list, selectedIndex);
      setTimeout(() => {
        const onDocDown = (ev: MouseEvent): void => {
          if (ev.target instanceof Node && wrap.contains(ev.target)) return;
          close();
        };
        document.addEventListener('mousedown', onDocDown, true);
        detachDocDown = () => document.removeEventListener('mousedown', onDocDown, true);
      }, 0);
    };

    button.addEventListener('click', () => {
      if (wrap.classList.contains('demo__rb-dd--open')) close();
      else open();
    });
    button.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        open();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        close();
      }
    });

    updateRibbonSelectDisplay(wrap, command);
    return wrap;
  };

  // Font / fill color button — an icon with a colored underline bar that opens
  // the shared color palette flyout (theme + standard colors,
  // "More Colors…" hands off to the native picker).
  const createRibbonColor = (command: RibbonCommand): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'demo__rb-color';
    wrap.dataset.ribbonCommand = command.id;

    const button = createRibbonControlButton({
      className: 'demo__rb-color__btn',
      title: command.title,
      ariaLabel: command.title,
      ariaHaspopup: 'true',
      ariaExpanded: false,
    });
    if (command.icon) {
      const icon = createRibbonIcon(command.icon);
      if (icon) {
        icon.classList.add('demo__rb-color__icon');
        button.appendChild(icon);
      }
    }
    const swatch = document.createElement('span');
    swatch.className = 'demo__rb-color__swatch';
    swatch.style.background = currentRibbonControlValue(command.id);
    button.append(swatch, makeSvg('0 0 12 12', RIBBON_CHEVRON_PATH, 'demo__rb-color__chev'));
    wrap.appendChild(button);

    // Hidden native picker, reached through the palette's "More Colors…" row.
    const native = document.createElement('input');
    native.type = 'color';
    native.className = 'demo__color-flyout__native';
    native.tabIndex = -1;
    native.setAttribute('aria-hidden', 'true');
    wrap.appendChild(native);

    let detachDocDown: (() => void) | null = null;
    const close = (): void => {
      wrap.classList.remove('demo__rb-color--open');
      button.setAttribute('aria-expanded', 'false');
      wrap.querySelector('.demo__color-flyout')?.remove();
      detachDocDown?.();
      detachDocDown = null;
    };
    const apply = (color: string): void => {
      applyRibbonControl(command.id, color);
      swatch.style.background = color;
    };
    native.addEventListener('input', () => apply(native.value));

    const open = (): void => {
      closeOpenRibbonDropdowns(wrap);
      wrap.classList.add('demo__rb-color--open');
      button.setAttribute('aria-expanded', 'true');
      const flyout = document.createElement('div');
      flyout.className = 'demo__color-flyout';
      const palette = createColorPalette({
        themeLabel: ribbonText.themeColors,
        standardLabel: ribbonText.standardColors,
        moreColorsLabel: ribbonText.moreColors,
        ariaLabel: command.title,
        value: currentRibbonControlValue(command.id),
        automatic:
          command.id === 'fontColor' ? { label: ribbonText.automatic, color: '#000000' } : null,
        onPick: (color) => {
          apply(color);
          close();
          button.focus({ preventScroll: true });
        },
        onMoreColors: () => {
          close();
          native.value = currentRibbonControlValue(command.id);
          native.click();
        },
      });
      flyout.appendChild(palette.el);
      const anchorRect = button.getBoundingClientRect();
      flyout.style.left = `${Math.round(anchorRect.left)}px`;
      flyout.style.top = `${Math.round(anchorRect.bottom + 3)}px`;
      flyout.addEventListener('keydown', (event) => {
        if (event.key !== 'Escape') return;
        event.preventDefault();
        close();
        button.focus({ preventScroll: true });
      });
      wrap.appendChild(flyout);
      palette.focus();
      setTimeout(() => {
        const onDocDown = (ev: MouseEvent): void => {
          if (ev.target instanceof Node && wrap.contains(ev.target)) return;
          close();
        };
        document.addEventListener('mousedown', onDocDown, true);
        detachDocDown = () => document.removeEventListener('mousedown', onDocDown, true);
      }, 0);
    };

    button.addEventListener('click', () => {
      if (wrap.classList.contains('demo__rb-color--open')) close();
      else open();
    });
    button.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        open();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        close();
      }
    });
    return wrap;
  };

  const ribbonSelectLabel = (wrap: HTMLElement, current: string): string => {
    if (wrap.dataset.ribbonSelect === 'sheetViewSelect') {
      if (current === 'current') return ribbonText.currentView;
      const state = getInst()?.store.getState();
      return state?.sheetViews.views.find((view) => view.id === current)?.name ?? current;
    }
    try {
      const options = JSON.parse(wrap.dataset.ribbonOptions ?? '[]') as {
        value: string;
        label: string;
      }[];
      const label = options.find((option) => option.value === current)?.label;
      if (label) return label;
      if (wrap.dataset.ribbonSelect === 'scalePercent') return `${current}%`;
      if (
        wrap.dataset.ribbonSelect === 'scaleWidth' ||
        wrap.dataset.ribbonSelect === 'scaleHeight'
      ) {
        if (current === '0') return pageScaleText.automatic;
        return `${current} ${current === '1' ? pageScaleText.page : pageScaleText.pages}`;
      }
      return current;
    } catch {
      return current;
    }
  };

  return {
    makeSvg,
    createRibbonSelect,
    createRibbonColor,
    closeOpenRibbonDropdowns,
    updateRibbonSelectDisplay,
    ribbonSelectLabel,
    RIBBON_CHEVRON_PATH,
  };
};
