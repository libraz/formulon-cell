import {
  type SpreadsheetCompatibilityId,
  type SpreadsheetCompatibilityStatus,
  summarizeSpreadsheetCompatibility,
} from '../engine/compatibility.js';
import {
  listWorkbookObjects,
  summarizePassthroughs,
  summarizePivotTables,
  summarizeTables,
  WORKBOOK_OBJECT_KINDS,
  workbookObjectKindCounts,
} from '../engine/passthrough-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

type WorkbookObjectsStrings = Strings['workbookObjects'];

export interface SpreadsheetCompatibilityReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

const COMPATIBILITY_LABEL_KEYS: Record<
  SpreadsheetCompatibilityId,
  keyof WorkbookObjectsStrings['compatibilityLabels']
> = {
  'cell-formatting': 'cellFormatting',
  'conditional-formatting': 'conditionalFormatting',
  'data-validation': 'dataValidation',
  hyperlinks: 'hyperlinks',
  comments: 'comments',
  'defined-names': 'definedNames',
  'sheet-protection': 'sheetProtection',
  'sheet-views': 'sheetViews',
  'loaded-tables': 'loadedTables',
  'format-as-table': 'formatAsTable',
  'pivot-layouts': 'pivotLayouts',
  'pivot-authoring': 'pivotAuthoring',
  'session-charts': 'sessionCharts',
  'charts-drawings': 'chartsDrawings',
  'chart-authoring': 'chartAuthoring',
  'external-links': 'externalLinks',
};

const STATUS_LABEL_KEYS: Record<
  SpreadsheetCompatibilityStatus,
  keyof Pick<WorkbookObjectsStrings, 'writable' | 'readOnly' | 'sessionOnly' | 'unsupported'>
> = {
  writable: 'writable',
  'read-only': 'readOnly',
  session: 'sessionOnly',
  unsupported: 'unsupported',
};

export const spreadsheetCompatibilityLabel = (
  id: SpreadsheetCompatibilityId,
  strings: WorkbookObjectsStrings,
): string => strings.compatibilityLabels[COMPATIBILITY_LABEL_KEYS[id]];

export const spreadsheetCompatibilityDetail = (
  id: SpreadsheetCompatibilityId,
  strings: WorkbookObjectsStrings,
): string => strings.compatibilityDetails[COMPATIBILITY_LABEL_KEYS[id]];

export const spreadsheetCompatibilityStatusLabel = (
  status: SpreadsheetCompatibilityStatus,
  strings: WorkbookObjectsStrings,
): string => strings[STATUS_LABEL_KEYS[status]];

/** Build the flat severity/label/detail list rendered by the React, Vue, and
 *  playground "inspect workbook" backstage actions. Three call sites used to
 *  carry ~100 lines of switch-case duplication each; this helper is the single
 *  source of truth for that mapping. */
export const buildSpreadsheetCompatibilityReport = (
  wb: WorkbookHandle,
  strings: WorkbookObjectsStrings,
): SpreadsheetCompatibilityReportItem[] => {
  const summary = summarizeSpreadsheetCompatibility(wb);
  const banner: SpreadsheetCompatibilityReportItem = {
    severity: 'info',
    label: strings.compatibility,
    detail:
      `${strings.writable} ${summary.byStatus.writable}, ` +
      `${strings.readOnly} ${summary.byStatus['read-only']}, ` +
      `${strings.sessionOnly} ${summary.byStatus.session}, ` +
      `${strings.unsupported} ${summary.byStatus.unsupported}`,
  };
  const rows = summary.items.map<SpreadsheetCompatibilityReportItem>((entry) => {
    const detail = spreadsheetCompatibilityDetail(entry.id, strings);
    return {
      severity: entry.status === 'unsupported' || entry.status === 'read-only' ? 'warning' : 'info',
      label: `${spreadsheetCompatibilityLabel(entry.id, strings)} · ${spreadsheetCompatibilityStatusLabel(entry.status, strings)}`,
      detail: entry.count === undefined ? detail : `${detail} (${entry.count})`,
    };
  });
  return [banner, ...rows];
};

export interface WorkbookObjectsPanelDeps {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings?: Strings;
}

export interface WorkbookObjectsPanelHandle {
  open(): void;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

const compatibilityLabelKey = (
  id: SpreadsheetCompatibilityId,
): keyof WorkbookObjectsStrings['compatibilityLabels'] => COMPATIBILITY_LABEL_KEYS[id];

export function attachWorkbookObjectsPanel(
  deps: WorkbookObjectsPanelDeps,
): WorkbookObjectsPanelHandle {
  const { host } = deps;
  let wb = deps.wb;
  let strings = deps.strings ?? defaultStrings;
  let open = false;
  let restoreFocusEl: HTMLElement | null = null;

  const root = document.createElement('div');
  root.className = 'fc-objects';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'false');
  root.hidden = true;
  root.tabIndex = -1;
  host.appendChild(root);
  inheritHostTokens(host, root);

  const close = (restoreFocus = false): void => {
    const wasOpen = open;
    open = false;
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    root.hidden = true;
    if (
      wasOpen &&
      restoreFocus &&
      focusTarget &&
      (root.contains(document.activeElement) || document.activeElement === document.body)
    ) {
      focusTarget.focus({ preventScroll: true });
    }
  };

  const item = (label: string, value: string | number): HTMLDivElement => {
    const row = document.createElement('div');
    row.className = 'fc-objects__row';
    const k = document.createElement('span');
    k.className = 'fc-objects__key';
    k.textContent = label;
    const v = document.createElement('span');
    v.className = 'fc-objects__value';
    v.textContent = String(value);
    row.append(k, v);
    return row;
  };

  const render = (): void => {
    const t = strings.workbookObjects;
    const objects = listWorkbookObjects(wb);
    const passthroughs = summarizePassthroughs(wb);
    const tables = summarizeTables(wb);
    const pivots = summarizePivotTables(wb);
    const support = summarizeSpreadsheetCompatibility(wb);
    root.replaceChildren();
    root.setAttribute('aria-label', t.title);

    const header = document.createElement('div');
    header.className = 'fc-objects__header';
    const title = document.createElement('div');
    title.className = 'fc-objects__title';
    title.textContent = t.title;
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'fc-objects__close';
    closeBtn.textContent = '×';
    closeBtn.setAttribute('aria-label', t.close);
    closeBtn.addEventListener('click', () => close(false));
    header.append(title, closeBtn);
    root.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-objects__body';
    const summary = document.createElement('section');
    summary.className = 'fc-objects__section';
    summary.append(
      item(t.preservedParts, passthroughs.count),
      item(t.tables, tables.count),
      item(t.pivotTables, pivots.count),
      item(t.writable, support.byStatus.writable),
      item(t.readOnly, support.byStatus['read-only']),
      item(t.sessionOnly, support.byStatus.session),
      item(t.unsupported, support.byStatus.unsupported),
      item(t.noteLabel, t.readOnlyNote),
    );
    body.appendChild(summary);

    const supportSection = document.createElement('section');
    supportSection.className = 'fc-objects__section';
    const supportHeading = document.createElement('div');
    supportHeading.className = 'fc-objects__heading';
    supportHeading.textContent = t.compatibility;
    supportSection.appendChild(supportHeading);
    const supportList = document.createElement('ul');
    supportList.className = 'fc-objects__paths';
    for (const entry of support.items) {
      const li = document.createElement('li');
      li.textContent = [
        t.compatibilityLabels[compatibilityLabelKey(entry.id)],
        t[
          entry.status === 'read-only'
            ? 'readOnly'
            : entry.status === 'session'
              ? 'sessionOnly'
              : entry.status
        ],
        entry.count === undefined ? '' : `${entry.count}`,
      ]
        .filter(Boolean)
        .join(' · ');
      supportList.appendChild(li);
    }
    supportSection.appendChild(supportList);
    body.appendChild(supportSection);

    const objectCounts = workbookObjectKindCounts(objects);
    const cats = WORKBOOK_OBJECT_KINDS.filter((kind) => objectCounts[kind] > 0);
    if (cats.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.categories;
      section.appendChild(heading);
      for (const category of cats) {
        section.appendChild(item(t.kindLabels[category], objectCounts[category]));
      }
      body.appendChild(section);
    }

    if (tables.names.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.tableNames;
      section.appendChild(heading);
      const list = document.createElement('div');
      list.className = 'fc-objects__list';
      list.textContent = tables.names.join(', ');
      section.appendChild(list);
      body.appendChild(section);
    }

    if (tables.items.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.tableDetails;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const table of tables.items) {
        const li = document.createElement('li');
        const name = table.displayName || table.name;
        const cols = table.columns.length;
        li.textContent = [
          name,
          `${t.sheet} ${table.sheetIndex + 1}`,
          table.ref,
          `${cols} ${cols === 1 ? t.columnSingular : t.columnPlural}`,
        ].join(' · ');
        list.appendChild(li);
      }
      section.appendChild(list);
      body.appendChild(section);
    }

    if (pivots.items.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.pivotDetails;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const pivot of pivots.items) {
        const li = document.createElement('li');
        const fields = pivot.fields.length > 0 ? ` · ${pivot.fields.join(', ')}` : '';
        li.textContent = [
          `${t.pivot} ${pivot.pivotIndex + 1}`,
          `${t.sheet} ${pivot.sheetIndex + 1}`,
          `R${pivot.top + 1}C${pivot.left + 1}`,
          `${pivot.rows} x ${pivot.cols}`,
          `${pivot.cells} ${t.cells}${fields}`,
        ].join(' · ');
        list.appendChild(li);
      }
      section.appendChild(list);
      body.appendChild(section);
    }

    if (objects.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = t.paths;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const object of objects.slice(0, 32)) {
        const li = document.createElement('li');
        li.textContent = `${t.kindLabels[object.kind]} · ${object.path}`;
        li.title = object.path;
        list.appendChild(li);
      }
      section.appendChild(list);
      body.appendChild(section);
    }

    if (passthroughs.count === 0 && tables.count === 0 && pivots.count === 0) {
      const empty = document.createElement('div');
      empty.className = 'fc-objects__empty';
      empty.textContent = t.empty;
      body.appendChild(empty);
    }
    root.appendChild(body);
  };

  const refresh = (): void => {
    if (open) render();
  };

  const openPanel = (): void => {
    render();
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.hidden = false;
    open = true;
    root.focus({ preventScroll: true });
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') close(true);
  };
  root.addEventListener('keydown', onKey);

  return {
    open: openPanel,
    close,
    refresh,
    setStrings(next) {
      strings = next;
      refresh();
    },
    bindWorkbook(next) {
      wb = next;
      refresh();
    },
    detach() {
      root.removeEventListener('keydown', onKey);
      root.remove();
    },
  };
}
