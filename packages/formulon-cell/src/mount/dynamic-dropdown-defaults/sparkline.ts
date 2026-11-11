import { type Color, SparklineAxisType } from '@libraz/formulon';
import { isMacChrome } from '../../commands/chrome-platform.js';
import type { ObjectActionContext } from '../../commands/object-actions.js';
import { isSheetProtected } from '../../commands/protection.js';
import { membersOfGroup } from '../../commands/sparkline.js';
import {
  activeSparklineGroup,
  clearSelectedSparklineGroupsAction,
  clearSelectedSparklinesAction,
  patchSelectedSparklineGroupsAction,
  type SparklineColorPart,
  setSparklineAxisBoundAction,
  setSparklineColorAction,
  setSparklineSourceAction,
  switchedSparklineSources,
  switchSparklineRowColumnAction,
} from '../../commands/sparkline-design.js';
import { unionSparklineSources } from '../../commands/sparkline-layout.js';
import {
  applySparklineStylePresetAction,
  SPARKLINE_STYLE_PRESETS,
  sparklineStylePresetIndex,
} from '../../commands/sparkline-presets.js';
import { sparklineSourceText } from '../../commands/sparkline-ranges.js';
import { syncColorPaletteTheme, themeColorColumns } from '../../components/color-palette.js';
import { addrKey, formatA1Range } from '../../engine/address.js';
import { cssOfColorSpec } from '../../engine/theme-color.js';
import type { SpreadsheetInstance } from '../../mount.js';
import { showNumberPrompt } from '../../toolbar/dialogs/prompt.js';
import {
  showSparklineHiddenEmptyDialog,
  showSparklineRangeDialog,
} from '../../toolbar/dialogs/sparkline.js';
import type {
  DynamicDropdownsCtx,
  SparklineMenuAction,
} from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { openMacSparklineDialog } from '../../toolbar/ribbon/mac/actions.js';
import {
  SPARKLINE_STYLE_SLOTS,
  sparklineStyleThumbnail,
} from '../../toolbar/ribbon/menus/sparkline.js';
import {
  SPARKLINE_AXIS_DISABLED_ITEMS,
  SPARKLINE_COMMANDS,
  SPARKLINE_MENU_IDS,
} from '../../toolbar/ribbon/sparkline-tabs.js';
import { setMenuControlDisabled } from './menu-feedback.js';

type SparklineDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  'applySparklineAction' | 'updateSparklineMenu'
>;

/** Bounds of a custom sparkline weight, in points. */
const MIN_WEIGHT_PT = 0.25;
const MAX_WEIGHT_PT = 1584;

const MARKER_PARTS: ReadonlySet<string> = new Set([
  'negative',
  'markers',
  'high',
  'low',
  'first',
  'last',
]);

export function createSparklineDropdownDefaults(
  instance: SpreadsheetInstance,
): SparklineDropdownDefaults {
  const ctxFor = (commandId: string): ObjectActionContext => ({
    store: instance.store,
    history: instance.history,
    origin: 'ribbon',
    commandId,
  });
  const strings = (): SpreadsheetInstance['i18n']['strings'] => instance.i18n.strings;
  const mac = (): boolean => isMacChrome(instance.store);
  const okCancel = () =>
    mac()
      ? { okLabel: strings().mac.sparkline.ok, cancelLabel: strings().mac.sparkline.cancel }
      : { okLabel: strings().hyperlinkDialog.ok, cancelLabel: strings().hyperlinkDialog.cancel };

  const promptNumber = (
    title: string,
    label: string,
    initial: number | undefined,
    bounds: { min?: number; max?: number } = {},
  ): Promise<number | null> =>
    showNumberPrompt({
      title,
      label,
      initial,
      ...bounds,
      ...okCancel(),
      invalidMessage: label,
      locale: () => instance.locale,
    });

  const promptRange = (title: string, label: string, initial: string) =>
    showSparklineRangeDialog({
      host: instance.host,
      store: instance.store,
      getWb: () => instance.workbook,
      title,
      label,
      initial,
      ...okCancel(),
      invalidMessage: label,
      pickerLabel: strings().pivotTableDialog.rangePickerSelect,
    });

  const editGroup = (): void => {
    const state = instance.store.getState();
    const group = activeSparklineGroup(state);
    if (!group) return;
    const members = membersOfGroup(state, group.id);
    const first = members[0];
    const last = members[members.length - 1];
    if (!first || !last) return;
    const rows = members.map(({ addr }) => addr.row);
    const cols = members.map(({ addr }) => addr.col);
    const location = formatA1Range({
      r0: Math.min(...rows),
      c0: Math.min(...cols),
      r1: Math.max(...rows),
      c1: Math.max(...cols),
    });
    openMacSparklineDialog(instance, {
      groupId: group.id,
      sheet: first.addr.sheet,
      data: unionSparklineSources(members.map(({ spec }) => spec.source)),
      location,
    });
  };

  const editSingle = async (): Promise<void> => {
    const state = instance.store.getState();
    const active = state.selection.active;
    const member = state.sparkline.sparklines.get(addrKey(active));
    if (!member) return;
    const t = strings().ribbon.sparklineDialog;
    const range = await promptRange(t.editSingleTitle, t.editSingleLabel, member.source);
    if (!range) return;
    setSparklineSourceAction(
      ctxFor(SPARKLINE_COMMANDS.editData),
      active,
      sparklineSourceText(range, active.sheet, instance.workbook),
    );
  };

  const hiddenEmpty = async (): Promise<void> => {
    const group = activeSparklineGroup(instance.store.getState());
    if (!group) return;
    const t = strings().ribbon.sparklineDialog;
    const result = await showSparklineHiddenEmptyDialog({
      host: instance.host,
      title: t.hiddenTitle,
      emptyCellsLabel: mac() ? strings().mac.sparkline.emptyCellsLabel : t.emptyCellsLabel,
      gaps: t.gaps,
      zero: t.zero,
      connect: t.connect,
      showHidden: mac() ? strings().mac.sparkline.showHiddenLabel : t.showHidden,
      ...okCancel(),
      initial: {
        emptyCellsAs: group.settings.emptyCellsAs,
        displayHidden: group.settings.displayHidden,
      },
    });
    if (result) patchSelectedSparklineGroupsAction(ctxFor(SPARKLINE_COMMANDS.editData), result);
  };

  const customWeight = async (): Promise<void> => {
    const group = activeSparklineGroup(instance.store.getState());
    if (!group) return;
    const m = strings().mac.sparkline;
    const t = strings().ribbon.sparklineDialog;
    const weight = await promptNumber(
      mac() ? m.weightTitle : t.weightTitle,
      mac() ? m.weightLabel : t.weightLabel,
      group.settings.lineWidthPt,
      { min: MIN_WEIGHT_PT, max: MAX_WEIGHT_PT },
    );
    if (weight !== null) {
      patchSelectedSparklineGroupsAction(ctxFor(SPARKLINE_COMMANDS.color), {
        lineWidthPt: weight,
      });
    }
  };

  const dateAxis = async (): Promise<void> => {
    const group = activeSparklineGroup(instance.store.getState());
    if (!group) return;
    const m = strings().mac.sparkline;
    const t = strings().ribbon.sparklineDialog;
    const current = group.settings.dateAxisRange;
    const range = await promptRange(
      mac() ? m.dateTitle : t.dateRangeTitle,
      mac() ? m.dateRangeLabel : t.dateRangeLabel,
      current
        ? formatA1Range({
            r0: current.rect.firstRow,
            c0: current.rect.firstCol,
            r1: current.rect.lastRow,
            c1: current.rect.lastCol,
          })
        : '',
    );
    if (!range) return;
    patchSelectedSparklineGroupsAction(ctxFor(SPARKLINE_COMMANDS.axis), {
      dateAxis: true,
      dateAxisRange: {
        sheet: range.sheet,
        rect: { firstRow: range.r0, firstCol: range.c0, lastRow: range.r1, lastCol: range.c1 },
      },
    });
  };

  const customBound = async (end: 'min' | 'max'): Promise<void> => {
    const group = activeSparklineGroup(instance.store.getState());
    if (!group) return;
    const m = strings().mac.sparkline;
    const t = strings().ribbon.sparklineDialog;
    const current = end === 'min' ? group.settings.manualMin : group.settings.manualMax;
    const value = await promptNumber(
      mac() ? m.customValueTitle : t.axisValueTitle,
      mac() ? m.customValueLabel : end === 'min' ? t.axisMinLabel : t.axisMaxLabel,
      current ?? undefined,
    );
    if (value !== null) {
      setSparklineAxisBoundAction(ctxFor(SPARKLINE_COMMANDS.axis), end, 'custom', value);
    }
  };

  const axis = (value: string): void | Promise<void> => {
    const ctx = ctxFor(SPARKLINE_COMMANDS.axis);
    if (value === 'general') {
      patchSelectedSparklineGroupsAction(ctx, { dateAxis: false });
    } else if (value === 'date') {
      return dateAxis();
    } else if (value === 'show') {
      const group = activeSparklineGroup(instance.store.getState());
      if (group)
        patchSelectedSparklineGroupsAction(ctx, { displayXAxis: !group.settings.displayXAxis });
    } else {
      const [end, mode] = value.split('-');
      if ((end !== 'min' && end !== 'max') || !mode) return;
      if (mode === 'custom') return customBound(end);
      if (mode === 'individual' || mode === 'group') {
        setSparklineAxisBoundAction(ctx, end, mode);
      }
    }
  };

  return {
    applySparklineAction(kind: SparklineMenuAction, value: string) {
      if (kind === 'style') {
        const index = SPARKLINE_STYLE_SLOTS.findIndex((slot) => slot.id === value);
        if (index >= 0)
          applySparklineStylePresetAction(ctxFor(SPARKLINE_COMMANDS.styleGallery), index);
      } else if (kind === 'color') {
        setSparklineColorAction(ctxFor(SPARKLINE_COMMANDS.color), 'series', value);
      } else if (kind === 'weight') {
        if (value === 'custom') return customWeight();
        const weight = Number(value);
        if (Number.isFinite(weight)) {
          patchSelectedSparklineGroupsAction(ctxFor(SPARKLINE_COMMANDS.color), {
            lineWidthPt: weight,
          });
        }
      } else if (kind === 'markerColor') {
        const split = value.indexOf(':');
        const part = value.slice(0, split);
        if (split > 0 && MARKER_PARTS.has(part)) {
          setSparklineColorAction(
            ctxFor(SPARKLINE_COMMANDS.markerColor),
            part as SparklineColorPart,
            value.slice(split + 1),
          );
        }
      } else if (kind === 'edit') {
        if (value === 'group') editGroup();
        else if (value === 'single') return editSingle();
        else if (value === 'hidden-empty') return hiddenEmpty();
        else if (value === 'switch')
          switchSparklineRowColumnAction(ctxFor(SPARKLINE_COMMANDS.editData));
      } else if (kind === 'axis') {
        return axis(value);
      } else if (value === 'groups') {
        clearSelectedSparklineGroupsAction(ctxFor(SPARKLINE_COMMANDS.clear));
      } else {
        clearSelectedSparklinesAction(ctxFor(SPARKLINE_COMMANDS.clear));
      }
    },

    updateSparklineMenu(menu) {
      const state = instance.store.getState();
      const group = activeSparklineGroup(state);
      const protectedSheet = isSheetProtected(state, state.selection.active.sheet);
      const inert = protectedSheet || group === null;
      const settings = group?.settings;
      const buttons = menu.querySelectorAll<HTMLButtonElement>('button');
      for (const button of buttons) setMenuControlDisabled(button, inert);
      if (menu.id === SPARKLINE_MENU_IDS.editData) {
        const sw = menu.querySelector<HTMLButtonElement>('[data-sparkline-edit="switch"]');
        if (sw) setMenuControlDisabled(sw, inert || switchedSparklineSources(state) === null);
      }
      if (menu.id === SPARKLINE_MENU_IDS.color || menu.id === SPARKLINE_MENU_IDS.markerColor) {
        syncColorPaletteTheme(menu, themeColorColumns(instance.workbook.getTheme()?.colors));
      }
      if (menu.id === SPARKLINE_MENU_IDS.color) {
        for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-sparkline-weight]')) {
          const value = button.dataset.sparklineWeight ?? '';
          const checked = !!settings && Number(value) === settings.lineWidthPt;
          button.setAttribute('aria-checked', String(checked));
          button.classList.toggle('fc-tb__menu-item--active', checked);
        }
      }
      if (menu.id === SPARKLINE_MENU_IDS.axis && settings) {
        const checked: Record<string, boolean> = {
          general: !settings.dateAxis,
          date: settings.dateAxis,
          show: settings.displayXAxis,
          'min-individual': settings.minAxisType === SparklineAxisType.Individual,
          'min-group': settings.minAxisType === SparklineAxisType.Group,
          'min-custom': settings.minAxisType === SparklineAxisType.Custom,
          'max-individual': settings.maxAxisType === SparklineAxisType.Individual,
          'max-group': settings.maxAxisType === SparklineAxisType.Group,
          'max-custom': settings.maxAxisType === SparklineAxisType.Custom,
        };
        for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-sparkline-axis]')) {
          const value = button.dataset.sparklineAxis ?? '';
          button.setAttribute('aria-checked', String(checked[value] === true));
          button.classList.toggle('fc-tb__menu-item--active', checked[value] === true);
          // The engine models no right-to-left plot, so that entry never runs.
          if (SPARKLINE_AXIS_DISABLED_ITEMS.includes(value)) setMenuControlDisabled(button, true);
        }
      }
      if (menu.id === SPARKLINE_MENU_IDS.style) {
        const current = sparklineStylePresetIndex(settings);
        const css = (color: Color): string => cssOfColorSpec(instance.workbook, color) ?? '';
        SPARKLINE_STYLE_SLOTS.forEach((slot, index) => {
          const swatch = menu.querySelector<HTMLButtonElement>(
            `[data-sparkline-style="${slot.id}"]`,
          );
          const preset = SPARKLINE_STYLE_PRESETS[index];
          if (!swatch || !preset) return;
          swatch.replaceChildren(sparklineStyleThumbnail(preset, css));
          swatch.setAttribute('aria-checked', String(index === current));
          swatch.classList.toggle('fc-tb__menu-item--active', index === current);
        });
      }
    },
  };
}
