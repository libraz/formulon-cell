import {
  ACCESSIBILITY_SECTIONS,
  type AccessibilityFinding,
  type AccessibilityFindings,
  type AccessibilityRuleId,
  type AccessibilitySectionId,
  checkAccessibility,
  hasAccessibilityIssues,
  suggestedTextColors,
} from '../commands/accessibility-check.js';
import { setFontColor } from '../commands/format.js';
import type { History } from '../commands/history.js';
import { recordFormatChange } from '../commands/slice-history.js';
import { toggleTableOption } from '../commands/table-design.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { MacAccessibilityPaneStrings } from '../i18n/strings/mac/accessibility-pane.js';
import type { Strings } from '../i18n/strings.js';
import { sameRange } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogButton } from './dialog-shell.js';
import { syncTaskpaneDock } from './taskpane-dock.js';

type FormatDialogTab = 'font' | 'fill';

export interface AccessibilityPaneDeps {
  /** The `.fc-host` element; the pane docks in its task-pane dock. */
  host: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle;
  getStrings: () => Strings;
  history?: History | null;
  /** Open Format Cells on `tab` over the current selection. */
  openFormatDialog?: (tab: FormatDialogTab) => void;
}

export interface AccessibilityPaneHandle {
  open(): void;
  close(): void;
  toggle(): void;
  isOpen(): boolean;
  /** The current findings, re-run on demand. */
  findings(): AccessibilityFindings;
  detach(): void;
}

const fill = (template: string, values: Record<string, string | number>): string =>
  template.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in values ? String(values[key]) : match,
  );

const sectionOf = (rule: AccessibilityRuleId): AccessibilitySectionId =>
  ACCESSIBILITY_SECTIONS.find((section) => section.rules.includes(rule))?.id ?? 'colorContrast';

/**
 * The desktop's accessibility checker as a docked task pane: a summary card,
 * the rule rows grouped by section with a count or a check each, and a
 * drill-in card per rule that selects the offending cells and offers fixes.
 */
export function attachAccessibilityPane(deps: AccessibilityPaneDeps): AccessibilityPaneHandle {
  const { host, store, getWb } = deps;
  const labels = (): MacAccessibilityPaneStrings => deps.getStrings().macAccessibilityPane;
  const dock = (): HTMLElement | null => host.querySelector<HTMLElement>('.fc-host__taskpane-dock');

  const root = document.createElement('div');
  root.className = 'fc-a11ypane';
  root.setAttribute('role', 'region');
  root.hidden = true;

  const header = document.createElement('div');
  header.className = 'fc-a11ypane__header';
  const title = document.createElement('span');
  title.className = 'fc-a11ypane__title';
  const closeBtn = appendDialogButton(header, { label: '', baseClass: 'fc-a11ypane__close' });
  closeBtn.append(createExcelRibbonSvg('paneClose', 'fc-pane-close__glyph') ?? '');
  header.prepend(title);
  const body = document.createElement('div');
  body.className = 'fc-a11ypane__body';
  root.append(header, body);
  (dock() ?? host).appendChild(root);

  let current: AccessibilityFindings = checkAccessibility(store.getState(), getWb());
  let drill: AccessibilityRuleId | null = null;

  const select = (range: Range): void => {
    if (store.getState().data.sheetIndex !== range.sheet) {
      mutators.setSheetIndex(store, range.sheet);
      mutators.replaceCells(store, getWb().cells(range.sheet));
    }
    mutators.setActive(store, { sheet: range.sheet, row: range.r0, col: range.c0 });
    mutators.extendRangeTo(store, { sheet: range.sheet, row: range.r1, col: range.c1 });
  };

  const button = (parent: HTMLElement, className: string, text = ''): HTMLButtonElement => {
    const el = appendDialogButton(parent, { label: text, baseClass: className });
    return el;
  };

  const renderOverview = (): void => {
    const t = labels();
    const clean = !hasAccessibilityIssues(current);
    const summary = document.createElement('div');
    summary.className = `fc-a11ypane__summary${clean ? ' fc-a11ypane__summary--clean' : ''}`;
    if (clean) {
      const text = document.createElement('p');
      text.className = 'fc-a11ypane__summary-text';
      text.textContent = `${t.cleanTitle}${t.cleanText}`;
      summary.appendChild(text);
    } else {
      const heading = document.createElement('h3');
      heading.className = 'fc-a11ypane__summary-title';
      heading.textContent = t.issuesTitle;
      const text = document.createElement('p');
      text.className = 'fc-a11ypane__summary-text';
      text.textContent = t.issuesText;
      summary.append(heading, text);
    }
    body.appendChild(summary);
    for (const section of ACCESSIBILITY_SECTIONS) {
      const heading = document.createElement('h3');
      heading.className = 'fc-a11ypane__section';
      heading.textContent = t.sections[section.id];
      const group = document.createElement('div');
      group.className = 'fc-a11ypane__group';
      group.setAttribute('role', 'group');
      group.setAttribute('aria-label', t.sections[section.id]);
      for (const rule of section.rules) {
        const count = current[rule].length;
        const row = button(group, 'fc-a11ypane__rule');
        row.dataset.rule = rule;
        const name = document.createElement('span');
        name.className = 'fc-a11ypane__rule-name';
        name.textContent = t.rules[rule];
        const state = document.createElement('span');
        state.className = count > 0 ? 'fc-a11ypane__rule-count' : 'fc-a11ypane__rule-pass';
        state.textContent = count > 0 ? String(count) : '';
        const chevron = document.createElement('span');
        chevron.className = 'fc-a11ypane__chevron';
        chevron.setAttribute('aria-hidden', 'true');
        row.append(name, state, chevron);
        row.setAttribute(
          'aria-label',
          `${t.rules[rule]}, ${count > 0 ? fill(t.issueCount, { count }) : t.noIssues}`,
        );
        row.addEventListener('click', () => {
          const first = current[rule][0];
          if (!first) return;
          drill = rule;
          render();
          select(first.range);
        });
      }
      body.append(heading, group);
    }
    const help = button(body, 'fc-a11ypane__help', '?');
    help.setAttribute('aria-label', t.help);
    projectDisabledState(help, true, null);
  };

  const renderDrill = (rule: AccessibilityRuleId, finding: AccessibilityFinding): void => {
    const t = labels();
    const top = document.createElement('div');
    top.className = 'fc-a11ypane__drill-head';
    const back = button(top, 'fc-a11ypane__back');
    back.setAttribute('aria-label', t.back);
    back.addEventListener('click', () => {
      drill = null;
      render();
    });
    const counter = document.createElement('h3');
    counter.className = 'fc-a11ypane__counter';
    counter.textContent = fill(t.counter, {
      index: 1,
      total: current[rule].length,
      section: t.sections[sectionOf(rule)],
    });
    top.appendChild(counter);
    const card = document.createElement('div');
    card.className = 'fc-a11ypane__card';
    const heading = document.createElement('h4');
    heading.className = 'fc-a11ypane__card-title';
    const text = document.createElement('p');
    text.className = 'fc-a11ypane__card-text';
    const tryFix = document.createElement('h5');
    tryFix.className = 'fc-a11ypane__try';
    tryFix.textContent = t.tryFix;
    card.append(heading, text, tryFix);
    if (rule === 'contrast') {
      heading.textContent = t.contrastTitle;
      text.textContent = t.contrastText;
      const swatches = document.createElement('div');
      swatches.className = 'fc-a11ypane__swatches';
      card.appendChild(swatches);
      for (const color of suggestedTextColors(
        finding.foreground ?? '#000000',
        finding.background ?? '#ffffff',
      )) {
        const swatch = button(swatches, 'fc-a11ypane__swatch');
        swatch.style.background = color;
        swatch.setAttribute('aria-label', fill(t.contrastSwatch, { color }));
        swatch.addEventListener('click', () => {
          recordFormatChange(deps.history ?? null, store, () =>
            setFontColor(store.getState(), store, color),
          );
        });
      }
      for (const [label, tab] of [
        [t.moreFontColors, 'font'],
        [t.fillColor, 'fill'],
        [t.formatCellFont, 'font'],
      ] as const) {
        const fix = button(card, 'fc-a11ypane__fix', label);
        projectDisabledState(fix, !deps.openFormatDialog, null);
        fix.addEventListener('click', () => deps.openFormatDialog?.(tab));
      }
    } else if (rule === 'tableHeader') {
      heading.textContent = t.tableHeaderTitle;
      text.textContent = t.tableHeaderText;
      const fix = button(card, 'fc-a11ypane__fix', t.insertHeaderRow);
      fix.addEventListener('click', () => {
        toggleTableOption(store, deps.history ?? null, getWb(), 'showHeader');
      });
    }
    const help = button(card, 'fc-a11ypane__help', '?');
    help.setAttribute('aria-label', t.help);
    projectDisabledState(help, true, null);
    body.append(top, card);
  };

  function render(): void {
    const t = labels();
    title.textContent = t.title;
    root.setAttribute('aria-label', t.title);
    closeBtn.setAttribute('aria-label', t.close);
    closeBtn.title = t.close;
    body.replaceChildren();
    const finding = drill ? current[drill][0] : undefined;
    if (drill && finding) renderDrill(drill, finding);
    else {
      drill = null;
      renderOverview();
    }
  }

  const rerun = (): void => {
    const previous = drill ? current[drill][0] : undefined;
    current = checkAccessibility(store.getState(), getWb());
    const next = drill ? current[drill][0] : undefined;
    render();
    // Once a finding is fixed, the next one is selected, as on drill-in.
    if (next && previous && !sameRange(next.range, previous.range)) select(next.range);
  };

  let scheduled = false;
  const unsubscribe = store.subscribe(() => {
    if (root.hidden || scheduled) return;
    scheduled = true;
    queueMicrotask(() => {
      scheduled = false;
      if (!root.hidden) rerun();
    });
  });

  const syncDock = (): void => {
    const target = dock();
    if (target) syncTaskpaneDock(target);
  };

  const open = (): void => {
    drill = null;
    current = checkAccessibility(store.getState(), getWb());
    render();
    root.hidden = false;
    syncDock();
  };
  const close = (): void => {
    root.hidden = true;
    syncDock();
    host.focus();
  };
  closeBtn.addEventListener('click', close);

  return {
    open,
    close,
    toggle: () => (root.hidden ? open() : close()),
    isOpen: () => !root.hidden,
    findings: () => checkAccessibility(store.getState(), getWb()),
    detach: () => {
      unsubscribe();
      root.remove();
      syncDock();
    },
  };
}
