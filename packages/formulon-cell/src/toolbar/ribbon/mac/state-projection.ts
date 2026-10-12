import { canExecuteBuiltIn } from '../../../commands/built-in-command-policy.js';
import { cellValueIsFormulaError } from '../../../commands/error-indicators.js';
import { addrKey, parseAddrKey } from '../../../engine/address.js';
import { getMacInk } from '../../../interact/mac-ink.js';
import type { SpreadsheetInstance } from '../../../mount/types.js';
import { projectDisabledState } from '../../menu-a11y.js';
import { indexRibbonButtons, projectActiveState } from '../../ribbon-active-state.js';
import { macRibbonLabelForCommand } from './labels.js';
import { MAC_CONTEXT_LABELS } from './labels-data.js';
import { toolbarLangForLocale } from './locale.js';

const MAC_ACTIVE_STATE_KEYS: Readonly<Record<string, keyof ReturnType<typeof projectActiveState>>> =
  {
    'mac.page.showGridlines': 'gridlinesVisible',
    'mac.page.printGridlines': 'printGridlines',
    'mac.page.showHeadings': 'headingsVisible',
    'mac.page.printHeadings': 'printHeadings',
    'mac.view.standard': 'workbookView',
    'mac.view.pageBreakPreview': 'workbookView',
    'mac.view.pageLayout': 'workbookView',
    'mac.view.gridlines': 'gridlinesVisible',
    'mac.view.headings': 'headingsVisible',
    'mac.view.zeros': 'zerosVisible',
    'mac.formulas.showFormulas': 'formulasVisible',
    'mac.data.filter': 'filterOn',
    'mac.review.protectSheet': 'protected',
    'mac.view.freeze': 'frozen',
  };

const MAC_POLICY_PROJECTED_COMMANDS = [
  'mac.page.theme',
  'mac.page.color',
  'mac.page.font',
  'mac.page.theme.office',
  'mac.page.theme.gray',
  'mac.page.color.office',
  'mac.page.color.gray',
  'mac.page.font.office',
  'mac.page.font.gray',
  'mac.insert.newComment',
  'mac.review.newComment',
  'mac.review.deleteComment',
  'mac.review.previousComment',
  'mac.review.nextComment',
  'mac.review.showComments',
  'mac.data.sortAsc',
  'mac.data.sortDesc',
  'mac.data.sortCustom',
  'mac.data.filter',
  'mac.data.clear',
  'mac.data.reapply',
  'mac.data.advancedFilter',
  'mac.formulas.removeArrows.all',
  'mac.formulas.removeArrows.precedents',
  'mac.formulas.removeArrows.dependents',
  'mac.formulas.errorCheck.run',
  'mac.formulas.errorCheck.trace',
  'mac.formulas.errorCheck.ignore',
  'mac.data.validation.settings',
  'mac.data.validation.circleInvalid',
  'mac.data.validation.clearCircles',
  'mac.data.validation.clearRules',
] as const;

const hasValidationInSelection = (instance: SpreadsheetInstance): boolean => {
  const { format, selection } = instance.store.getState();
  const { sheet, r0, r1, c0, c1 } = selection.range;
  const area = (r1 - r0 + 1) * (c1 - c0 + 1);
  if (area <= format.formats.size) {
    for (let row = r0; row <= r1; row += 1) {
      for (let col = c0; col <= c1; col += 1) {
        if (format.formats.get(addrKey({ sheet, row, col }))?.validation) return true;
      }
    }
    return false;
  }
  for (const [key, cellFormat] of format.formats) {
    if (!cellFormat.validation) continue;
    const addr = parseAddrKey(key);
    if (
      addr &&
      addr.sheet === sheet &&
      addr.row >= r0 &&
      addr.row <= r1 &&
      addr.col >= c0 &&
      addr.col <= c1
    ) {
      return true;
    }
  }
  return false;
};

const contextualDisabledReason = (instance: SpreadsheetInstance, id: string): string | null => {
  const state = instance.store.getState();
  const strings = instance.i18n.strings.ribbonMenu;
  const lang = toolbarLangForLocale(instance.i18n.locale);
  if (/^mac\.page\.(theme|color|font)(\.|$)/.test(id)) {
    if (!instance.workbook?.capabilities?.workbookTheme)
      return MAC_CONTEXT_LABELS.themeEngineUnavailable[lang];
    if (state.protection.protectedSheets.size > 0) return MAC_CONTEXT_LABELS.themeProtected[lang];
    if (instance.workbook.getTheme()?.source === 2) return MAC_CONTEXT_LABELS.themeUnreadable[lang];
  }
  if (
    id === 'mac.insert.newComment' ||
    (id.startsWith('mac.review.') &&
      ['newComment', 'deleteComment', 'previousComment', 'nextComment', 'showComments'].includes(
        id.slice('mac.review.'.length),
      ))
  ) {
    if (!instance.workbook?.capabilities?.threadedComments)
      return MAC_CONTEXT_LABELS.threadedCommentsUnavailable[lang];
    const roots =
      state.threadedComments.bySheet
        .get(state.data.sheetIndex)
        ?.filter((entry) => !entry.parentId) ?? [];
    if (
      id === 'mac.review.deleteComment' &&
      !roots.some(
        (entry) =>
          entry.row === state.selection.active.row && entry.col === state.selection.active.col,
      )
    )
      return strings.commentDeleteRequiresActive;
    if (
      ['mac.review.previousComment', 'mac.review.nextComment', 'mac.review.showComments'].includes(
        id,
      ) &&
      !roots.length
    )
      return strings.commentNone;
  }
  if (id === 'mac.data.clear' && !state.ui.filterRange) {
    return MAC_CONTEXT_LABELS.filterRangeRequired[lang];
  }
  if (id === 'mac.data.reapply' && state.ui.filterCriteria.length === 0) {
    return MAC_CONTEXT_LABELS.filterCriteriaRequired[lang];
  }

  if (id.startsWith('mac.formulas.removeArrows.')) {
    const traces = state.traces.items;
    const hasPrecedents = traces.some((trace) => trace.kind === 'precedent');
    const hasDependents = traces.some((trace) => trace.kind === 'dependent');
    if (id === 'mac.formulas.removeArrows.all' && !hasPrecedents && !hasDependents) {
      return strings.removeArrowsRequiresAny;
    }
    if (id === 'mac.formulas.removeArrows.precedents' && !hasPrecedents) {
      return strings.removePrecedentArrowsRequiresAny;
    }
    if (id === 'mac.formulas.removeArrows.dependents' && !hasDependents) {
      return strings.removeDependentArrowsRequiresAny;
    }
  }

  if (id === 'mac.formulas.errorCheck.trace' || id === 'mac.formulas.errorCheck.ignore') {
    const active = state.selection.active;
    const cell = state.data.cells.get(addrKey(active));
    if (!cell?.formula || !cellValueIsFormulaError(cell.value)) {
      return strings.traceErrorRequiresFormulaError;
    }
  }

  if (id === 'mac.data.validation.circleInvalid' || id === 'mac.data.validation.clearRules') {
    if (!hasValidationInSelection(instance)) return strings.validationRequiresRules;
  }
  if (
    id === 'mac.data.validation.clearCircles' &&
    state.errorIndicators.validationCircles.size === 0
  ) {
    return strings.validationClearCirclesRequiresAny;
  }
  return null;
};

/** Projects Mac alias buttons onto the same store-derived state as the
 *  generic ribbon. Hosts call this after an instance mutation; it is kept
 *  separate from the render pass so active controls do not require a full
 *  menu rebuild. */
export const projectMacRibbonState = (
  host: HTMLElement,
  instance: SpreadsheetInstance | null,
): void => {
  if (!instance) return;
  const lang = toolbarLangForLocale(instance.i18n.locale);
  const active = projectActiveState(instance);
  const ink = getMacInk(instance);
  const inkTool = ink?.getTool();
  const buttons = indexRibbonButtons(host);
  const drawingStates: Readonly<Record<string, boolean>> = {
    'mac.draw.toggle': ink?.isActive() ?? false,
    'mac.draw.eraser': inkTool === 'eraser',
    'mac.draw.penBlack': inkTool === 'pen-black',
    'mac.draw.penRed': inkTool === 'pen-red',
    'mac.draw.pencil': inkTool === 'pencil',
    'mac.draw.highlighter': inkTool === 'highlighter',
    'mac.draw.trackpad': ink?.isActive() === true && ink.getTrackpadMode(),
  };
  for (const [command, pressed] of Object.entries(drawingStates)) {
    for (const button of buttons.get(command) ?? []) {
      button.classList.toggle('fc-tb__rb--active', pressed);
      button.setAttribute('aria-pressed', String(pressed));
    }
  }

  for (const [command, key] of Object.entries(MAC_ACTIVE_STATE_KEYS)) {
    let pressed = Boolean(active[key]);
    if (
      command === 'mac.view.standard' ||
      command === 'mac.view.pageLayout' ||
      command === 'mac.view.pageBreakPreview'
    ) {
      const view = command.endsWith('standard')
        ? 'normal'
        : command.endsWith('pageLayout')
          ? 'pageLayout'
          : 'pageBreakPreview';
      pressed = active.workbookView === view;
    }
    for (const button of buttons.get(command) ?? []) {
      button.classList.toggle('fc-tb__rb--active', pressed);
      button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
    }
  }

  for (const id of MAC_POLICY_PROJECTED_COMMANDS) {
    const decision = canExecuteBuiltIn(instance.store, id, 'ribbon');
    const contextualReason = contextualDisabledReason(instance, id);
    const disabled = !decision.allowed || contextualReason !== null;
    const reason = decision.allowed
      ? contextualReason
      : instance.i18n.strings.backstage.commandUnavailable;
    for (const button of buttons.get(id) ?? []) {
      const titlePrefix =
        button.getAttribute('aria-label')?.trim() || macRibbonLabelForCommand(id, lang);
      projectDisabledState(button, disabled, reason, {
        datasetKey: 'ribbonDisabledReason',
        titlePrefix,
      });
    }
  }
};
