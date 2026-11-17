import { clamp } from '../charts/numeric.js';
import type { History } from '../commands/history.js';
import {
  canDeleteIllustration,
  canDuplicateIllustration,
  illustrationCopyRefusal,
  illustrationName,
} from '../commands/session-illustration.js';
import {
  cornerAdjForRadius,
  DEFAULT_SHAPE_COLOR,
  DEFAULT_SHAPE_LINE_WIDTH,
  MAX_SHAPE_LINE_WIDTH,
  shapeCornerRadius,
} from '../commands/shape-format.js';
import { summarizeSpreadsheetCompatibility } from '../engine/compatibility.js';
import {
  listWorkbookObjects,
  summarizePassthroughs,
  summarizePivotTables,
  summarizeTables,
  WORKBOOK_OBJECT_KINDS,
  workbookObjectKindCounts,
} from '../engine/passthrough-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  mutators,
  type SessionIllustration,
  type SessionShapeKind,
  type SpreadsheetStore,
} from '../store/store.js';
import { commandRefusalMessage } from '../toolbar/dialogs/command-refusal.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogIconButton,
  bindPaneEscape,
  markSurface,
  openerOf,
  returnFocus,
} from './dialog-shell.js';
import { createPivotNumberFormatPicker } from './pivot-field-settings-dialog.js';
import { openPivotFieldSettingsAt } from './pivot-field-settings-dialog-apply.js';
import { openPivotOptions } from './pivot-options-dialog.js';
import { renderedPosition } from './session-illustrations.js';
import { compatibilityLabelKey } from './spreadsheet-compatibility-report.js';
import { syncTaskpaneDock } from './taskpane-dock.js';
import { createWorkbookObjectsActionButton, pivotEditField } from './workbook-objects-dom.js';
import { type PivotEditorContext, renderPivotEditForm } from './workbook-objects-pivot-editor.js';
import {
  createPivotFieldPaneState,
  type PivotFieldPaneState,
  renderPivotFieldListPane,
} from './workbook-objects-pivot-field-pane.js';

export {
  buildFileLossReport,
  buildSpreadsheetCompatibilityReport,
  type SpreadsheetCompatibilityReportItem,
  spreadsheetCompatibilityDetail,
  spreadsheetCompatibilityLabel,
  spreadsheetCompatibilityStatusLabel,
} from './spreadsheet-compatibility-report.js';

export interface WorkbookObjectsPanelDeps {
  host: HTMLElement;
  wb: WorkbookHandle;
  strings?: Strings;
  /** Field list pane edits push their undo steps here. */
  history?: History | null;
  /** The host's task-pane dock beside the grid. The PivotTable Fields pane
   *  docks there; without one it floats over the right edge of `host`. */
  dock?: HTMLElement;
  /** Authorizes pivot edits; also lets the PivotTable Field dialog open Format Cells. */
  store: SpreadsheetStore;
  getLocale: () => string;
  onOpenPivotTableDialog?: () => void;
  onAfterPivotEdit?: () => void;
  onSelectSessionIllustration?: (id: string) => void;
  onDuplicateSessionIllustration?: (id: string) => void;
  onClearSessionIllustration?: (id: string) => void;
  onUpdateSessionIllustration?: (
    id: string,
    patch: Partial<Omit<SessionIllustration, 'id'>>,
  ) => void;
  listSessionIllustrations?: () => readonly SessionIllustration[];
  subscribeSessionObjects?: (listener: () => void) => () => void;
}

export interface WorkbookObjectsPanelHandle {
  open(): void;
  openPivotFieldList(sheetIndex: number, pivotIndex: number): boolean;
  /** Opens the PivotTable Field dialog on the field of the report cell at `addr`;
   *  false when the cell names no field. */
  openPivotFieldSettings(addr: Addr): boolean;
  /** Opens the PivotTable Options dialog on a report; false when there is no such report. */
  openPivotOptions(sheetIndex: number, pivotIndex: number): boolean;
  isPivotFieldListOpen(): boolean;
  /** Closes the field list and keeps it closed when the selection moves onto
   *  a report, until it is opened again explicitly. */
  hidePivotFieldList(): void;
  /** Whether Hide Field List is in effect. */
  isPivotFieldListHidden(): boolean;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

const SHAPE_KIND_LABEL: Record<
  SessionShapeKind,
  | 'shapeRectangle'
  | 'shapeRoundedRectangle'
  | 'shapeOval'
  | 'shapeTriangle'
  | 'shapeDiamond'
  | 'shapeLine'
  | 'shapeArrow'
> = {
  rectangle: 'shapeRectangle',
  'rounded-rectangle': 'shapeRoundedRectangle',
  oval: 'shapeOval',
  triangle: 'shapeTriangle',
  diamond: 'shapeDiamond',
  line: 'shapeLine',
  arrow: 'shapeArrow',
};

/** The kind of a drawing as the shapes gallery names it. */
const illustrationKind = (item: SessionIllustration, strings: Strings): string => {
  if (item.kind === 'image') return strings.ribbon.pictures;
  if (item.textBox) return strings.ribbonMenu.shapeTextBox;
  return item.shape ? strings.ribbonMenu[SHAPE_KIND_LABEL[item.shape]] : strings.ribbon.shapes;
};

export function attachWorkbookObjectsPanel(
  deps: WorkbookObjectsPanelDeps,
): WorkbookObjectsPanelHandle {
  const { host } = deps;
  let wb = deps.wb;
  let strings = deps.strings ?? defaultStrings;
  let open = false;
  let restoreFocusEl: HTMLElement | null = null;
  let activePivotEditKey = '';
  let activePivotFieldListKey = '';
  let pivotFieldListHidden = false;
  let pivotEditError = '';
  let pivotPaneState: PivotFieldPaneState = createPivotFieldPaneState();
  const numberFormatPicker = deps.store
    ? createPivotNumberFormatPicker({
        host,
        store: deps.store,
        getStrings: () => strings,
        getWb: () => wb,
        getLocale: deps.getLocale,
      })
    : null;

  const root = document.createElement('div');
  root.className = 'fc-objects';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'false');
  markSurface(root, 'floating');
  root.hidden = true;
  root.tabIndex = -1;
  host.appendChild(root);

  // The field list pane takes the dock (narrowing the grid); the general panel
  // floats over the host. The dock shows while anything in it is visible.
  const placeRoot = (docked: boolean): void => {
    const dock = deps.dock;
    if (!dock) return;
    const parent = docked ? dock : host;
    if (root.parentElement !== parent) parent.appendChild(root);
    syncTaskpaneDock(dock);
  };

  // The store mirrors the panel's general (selection pane) open state, so the
  // ribbon toggle follows every open and close path; a field list is not it.
  const syncOpenFlag = (): void => {
    if (deps.store) {
      mutators.setWorkbookObjectsOpen(deps.store, open && activePivotFieldListKey === '');
    }
  };

  const close = (restoreFocus = false): void => {
    const wasOpen = open;
    open = false;
    const focusTarget = restoreFocusEl;
    restoreFocusEl = null;
    root.hidden = true;
    placeRoot(false);
    syncOpenFlag();
    if (
      wasOpen &&
      restoreFocus &&
      (root.contains(document.activeElement) || document.activeElement === document.body)
    ) {
      returnFocus(focusTarget, root, host);
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

  const pivotAnchor = (pivot: {
    top: number;
    left: number;
    rows: number;
    cols: number;
    cells: number;
  }): string =>
    `R${pivot.top + 1}C${pivot.left + 1} · ${pivot.rows} x ${pivot.cols} · ${pivot.cells} ${strings.workbookObjects.cells}`;

  const fieldChips = (fields: readonly string[]): HTMLDivElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-objects__chips';
    for (const field of fields) {
      const chip = document.createElement('span');
      chip.className = 'fc-objects__chip';
      chip.textContent = field;
      wrap.appendChild(chip);
    }
    return wrap;
  };

  const pivotKey = (sheet: number, index: number): string => `${sheet}:${index}`;

  const appendPivotEditButtons = (
    actions: HTMLElement,
    pivot: { sheetIndex: number; pivotIndex: number },
  ): void => {
    const t = strings.workbookObjects;
    const key = pivotKey(pivot.sheetIndex, pivot.pivotIndex);
    const edit = createWorkbookObjectsActionButton(t.editPivotTable);
    edit.addEventListener('click', () => {
      activePivotEditKey = activePivotEditKey === key ? '' : key;
      activePivotFieldListKey = '';
      pivotEditError = '';
      render();
    });
    actions.appendChild(edit);
    const fieldList = createWorkbookObjectsActionButton(t.pivotFieldList);
    fieldList.addEventListener('click', () => {
      activePivotFieldListKey = activePivotFieldListKey === key ? '' : key;
      pivotPaneState = createPivotFieldPaneState();
      activePivotEditKey = '';
      pivotEditError = '';
      render();
    });
    actions.appendChild(fieldList);
  };

  const appendIllustrationActions = (
    li: HTMLLIElement,
    illustration: SessionIllustration,
  ): void => {
    const refusal = illustrationCopyRefusal(illustration);
    const reason = refusal ? commandRefusalMessage(refusal, strings) : null;
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions fc-objects__illustration-actions';
    const select = createWorkbookObjectsActionButton(strings.workbookObjects.objectSelect);
    select.addEventListener('click', () => deps.onSelectSessionIllustration?.(illustration.id));
    actions.appendChild(select);
    if (deps.onDuplicateSessionIllustration) {
      const duplicate = createWorkbookObjectsActionButton(strings.workbookObjects.objectDuplicate);
      projectDisabledState(duplicate, !canDuplicateIllustration(illustration), reason);
      duplicate.addEventListener('click', () =>
        deps.onDuplicateSessionIllustration?.(illustration.id),
      );
      actions.appendChild(duplicate);
    }
    if (deps.onClearSessionIllustration) {
      const remove = createWorkbookObjectsActionButton(strings.workbookObjects.objectDelete);
      projectDisabledState(remove, !canDeleteIllustration(illustration), reason);
      remove.addEventListener('click', () => deps.onClearSessionIllustration?.(illustration.id));
      actions.appendChild(remove);
    }
    li.appendChild(actions);
  };

  const renderIllustrationEditForm = (illustration: SessionIllustration): HTMLFormElement => {
    const t = strings.workbookObjects;
    const form = document.createElement('form');
    form.className = 'fc-objects__pivot-edit fc-objects__illustration-edit';
    const color = document.createElement('input');
    color.className = 'fc-objects__input fc-objects__color-input';
    color.type = 'color';
    color.value = illustration.color ?? DEFAULT_SHAPE_COLOR;
    const radius = document.createElement('input');
    radius.className = 'fc-objects__input';
    radius.type = 'number';
    radius.min = '0';
    radius.step = '1';
    const frame = renderedPosition(deps.store, illustration);
    radius.max = String(Math.floor(Math.min(frame.w, frame.h) / 2));
    radius.value = String(Math.round(shapeCornerRadius(illustration, frame.w, frame.h)));
    const initialRadius = radius.value;
    const lineWidth = document.createElement('input');
    lineWidth.className = 'fc-objects__input';
    lineWidth.type = 'number';
    lineWidth.min = '0';
    lineWidth.max = String(MAX_SHAPE_LINE_WIDTH);
    lineWidth.step = '0.25';
    lineWidth.value = String(illustration.lineWidth ?? DEFAULT_SHAPE_LINE_WIDTH);
    const opacity = document.createElement('input');
    opacity.className = 'fc-objects__input';
    opacity.type = 'range';
    opacity.min = '0';
    opacity.max = '1';
    opacity.step = '0.05';
    opacity.value = String(illustration.opacity ?? 1);
    // Only a rounded rectangle has a corner.
    const cornerless = illustration.shape !== 'rounded-rectangle';
    projectDisabledState(radius, cornerless, null);
    const locked = illustration.nativeLocked === true;
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions';
    const apply = createWorkbookObjectsActionButton(t.apply, { primary: true, type: 'submit' });
    actions.appendChild(apply);
    if (locked) {
      // A shape drawn only as its box has no style of its own.
      for (const input of [color, radius, lineWidth, opacity, apply]) {
        projectDisabledState(input, true, t.shapeLocked);
      }
    }
    // A paint the cell cannot state stays as the file holds it.
    const fillKept = illustration.nativeFillOther === true;
    const lineKept = illustration.nativeLineOther === true;
    if (fillKept) {
      for (const input of [color, opacity]) projectDisabledState(input, true, t.shapePaintKept);
    }
    if (lineKept) projectDisabledState(lineWidth, true, t.shapePaintKept);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      if (locked) return;
      deps.onUpdateSessionIllustration?.(illustration.id, {
        ...(fillKept ? {} : { color: color.value }),
        // An untouched corner keeps following the shape's size.
        ...(cornerless || radius.value === initialRadius
          ? {}
          : { cornerAdj: cornerAdjForRadius(Number(radius.value) || 0, frame.w, frame.h) }),
        ...(lineKept
          ? {}
          : { lineWidth: clamp(Number(lineWidth.value) || 0, 0, MAX_SHAPE_LINE_WIDTH) }),
        ...(fillKept ? {} : { opacity: clamp(Number(opacity.value) || 0, 0, 1) }),
      });
    });
    form.append(
      pivotEditField(t.shapeColor, color),
      pivotEditField(t.shapeRadius, radius),
      pivotEditField(t.shapeLineWidth, lineWidth),
      pivotEditField(t.shapeOpacity, opacity),
      actions,
    );
    return form;
  };

  const pivotEditorContext = (): PivotEditorContext => ({
    host,
    wb,
    strings,
    history: deps.history ?? null,
    store: deps.store,
    origin: 'ribbon',
    commandId: 'pivotFieldList',
    onChanged: () => deps.onAfterPivotEdit?.(),
    errorText: pivotEditError,
    setError: (message) => {
      pivotEditError = message;
    },
    rerender: () => render(),
    onRemoved: () => {
      activePivotEditKey = '';
      deps.onAfterPivotEdit?.();
    },
    onApplied: () => {
      activePivotEditKey = '';
      activePivotFieldListKey = '';
    },
  });

  const render = (): void => {
    const t = strings.workbookObjects;
    const objects = listWorkbookObjects(wb);
    const passthroughs = summarizePassthroughs(wb);
    const tables = summarizeTables(wb);
    const pivots = summarizePivotTables(wb);
    const illustrations = deps.listSessionIllustrations?.() ?? [];
    const support = summarizeSpreadsheetCompatibility(wb);
    const activeFieldListPivot = activePivotFieldListKey
      ? pivots.items.find(
          (pivot) => pivotKey(pivot.sheetIndex, pivot.pivotIndex) === activePivotFieldListKey,
        )
      : undefined;
    if (activePivotFieldListKey && !activeFieldListPivot) activePivotFieldListKey = '';
    syncOpenFlag();
    root.replaceChildren();
    root.className = `fc-objects${activeFieldListPivot ? ' fc-objects--taskpane' : ''}`;
    markSurface(root, activeFieldListPivot ? 'pane' : 'floating');
    if (open) placeRoot(activeFieldListPivot !== undefined);
    root.setAttribute('aria-label', activeFieldListPivot ? t.pivotFieldList : t.title);

    const header = document.createElement('div');
    header.className = 'fc-objects__header';
    const title = document.createElement('div');
    title.className = 'fc-objects__title';
    title.textContent = activeFieldListPivot ? t.pivotFieldList : t.title;
    const headerActions = document.createElement('div');
    headerActions.className = 'fc-objects__header-actions';
    const closeBtn = appendDialogIconButton(headerActions, {
      label: '',
      ariaLabel: t.close,
      baseClass: 'fc-objects__close',
    });
    closeBtn.addEventListener('click', () => close(false));
    header.append(title, headerActions);
    root.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-objects__body';
    if (activeFieldListPivot) {
      body.appendChild(
        renderPivotFieldListPane(
          {
            host,
            wb,
            strings,
            history: deps.history ?? null,
            store: deps.store,
            origin: 'ribbon',
            commandId: 'pivotFieldList',
            state: pivotPaneState,
            errorText: pivotEditError,
            setError: (message) => {
              pivotEditError = message;
            },
            rerender: refresh,
            ...(numberFormatPicker ? { pickNumberFormat: numberFormatPicker.pick } : {}),
            onChanged: () => {
              deps.onAfterPivotEdit?.();
              refresh();
            },
          },
          activeFieldListPivot,
        ),
      );
      root.appendChild(body);
      return;
    }
    const summary = document.createElement('section');
    summary.className = 'fc-objects__section';
    summary.append(
      item(t.preservedParts, passthroughs.count),
      item(t.tables, tables.count),
      item(t.pivotTables, pivots.count),
      item(strings.ribbon.illustrations, illustrations.length),
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
      for (const pivot of pivots.items) {
        const card = document.createElement('div');
        card.className = 'fc-objects__pivot-card';
        const titleRow = document.createElement('div');
        titleRow.className = 'fc-objects__pivot-title';
        const title = document.createElement('strong');
        title.textContent = `${t.pivot} ${pivot.pivotIndex + 1}`;
        const meta = document.createElement('span');
        meta.textContent = `${t.sheet} ${pivot.sheetIndex + 1} · ${pivotAnchor(pivot)}`;
        titleRow.append(title, meta);
        card.appendChild(titleRow);
        if (pivot.fields.length > 0) card.appendChild(fieldChips(pivot.fields));
        const canEditPivot = wb.capabilities.pivotTableMutate;
        if (deps.onOpenPivotTableDialog || canEditPivot) {
          const actions = document.createElement('div');
          actions.className = 'fc-objects__actions';
          if (canEditPivot) appendPivotEditButtons(actions, pivot);
          if (deps.onOpenPivotTableDialog) {
            const button = createWorkbookObjectsActionButton(t.createPivotTable);
            button.addEventListener('click', () => deps.onOpenPivotTableDialog?.());
            actions.appendChild(button);
          }
          card.appendChild(actions);
        }
        if (canEditPivot && activePivotEditKey === pivotKey(pivot.sheetIndex, pivot.pivotIndex)) {
          card.appendChild(renderPivotEditForm(pivotEditorContext(), pivot));
        }
        section.appendChild(card);
      }
      body.appendChild(section);
    }

    if (illustrations.length > 0) {
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = strings.ribbon.illustrations;
      section.appendChild(heading);
      const list = document.createElement('ul');
      list.className = 'fc-objects__paths';
      for (const illustration of illustrations) {
        const li = document.createElement('li');
        li.textContent = [
          illustrationName(illustrations, illustration, strings.sessionIllustrations),
          illustrationKind(illustration, strings),
          `${t.sheet} ${illustration.sheet + 1}`,
        ].join(' · ');
        if (illustration.src) li.title = illustration.src;
        appendIllustrationActions(li, illustration);
        if (illustration.kind === 'shape') li.appendChild(renderIllustrationEditForm(illustration));
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

    if (
      passthroughs.count === 0 &&
      tables.count === 0 &&
      pivots.count === 0 &&
      illustrations.length === 0
    ) {
      const empty = document.createElement('div');
      empty.className = 'fc-objects__empty';
      empty.textContent = t.empty;
      body.appendChild(empty);
    }
    root.appendChild(body);
  };

  // A re-render replaces every control, so focus returns to the control with the same tag,
  // class, label and data attributes, the same one of its kind; failing that, to the panel.
  const focusKeyOf = (el: Element): string =>
    [
      el.tagName,
      el.className,
      el.getAttribute('aria-label') ?? '',
      ...Array.from(el.attributes)
        .filter((attr) => attr.name.startsWith('data-pivot'))
        .map((attr) => `${attr.name}=${attr.value}`),
    ].join('|');
  const refresh = (): void => {
    if (!open) return;
    const focused = document.activeElement;
    const inside = focused instanceof HTMLElement && root.contains(focused) && focused !== root;
    const focusKey = inside ? focusKeyOf(focused) : null;
    const sameKind = (key: string): HTMLElement[] =>
      Array.from(root.querySelectorAll<HTMLElement>('*')).filter((el) => focusKeyOf(el) === key);
    const ordinal = focusKey ? sameKind(focusKey).indexOf(focused as HTMLElement) : -1;
    render();
    if (focusKey === null || root.contains(document.activeElement)) return;
    const target = sameKind(focusKey)[ordinal] ?? root;
    target.focus({ preventScroll: true });
  };
  const unsubscribeSessionObjects = deps.subscribeSessionObjects?.(refresh) ?? null;

  const openPanel = (): void => {
    render();
    // Reopening keeps the opener recorded the first time, not a control inside the surface.
    if (!root.contains(document.activeElement)) {
      restoreFocusEl = openerOf(document);
    }
    root.hidden = false;
    open = true;
    placeRoot(root.classList.contains('fc-objects--taskpane'));
    syncOpenFlag();
    root.focus({ preventScroll: true });
  };

  const openPivotFieldList = (sheetIndex: number, pivotIndex: number): boolean => {
    const key = pivotKey(sheetIndex, pivotIndex);
    if (
      !summarizePivotTables(wb).items.some(
        (pivot) => pivotKey(pivot.sheetIndex, pivot.pivotIndex) === key,
      )
    ) {
      return false;
    }
    activePivotEditKey = '';
    if (activePivotFieldListKey !== key) pivotPaneState = createPivotFieldPaneState();
    activePivotFieldListKey = key;
    pivotFieldListHidden = false;
    pivotEditError = '';
    openPanel();
    return true;
  };

  // Selection Pane shows the object list whichever pivot view the pane showed last.
  const openObjectList = (): void => {
    activePivotEditKey = '';
    activePivotFieldListKey = '';
    pivotEditError = '';
    openPanel();
  };

  // The object list floats and closes on Escape; the docked field list is a pane and stays open.
  const onKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape' || root.dataset.fcSurface !== 'floating') return;
    e.preventDefault();
    close(true);
  };
  root.addEventListener('keydown', onKey);
  const detachPaneEscape = bindPaneEscape(root, host);

  return {
    open: openObjectList,
    openPivotFieldList,
    openPivotFieldSettings: (addr) =>
      openPivotFieldSettingsAt(
        {
          host,
          strings,
          wb,
          history: deps.history ?? null,
          store: deps.store,
          origin: 'ribbon',
          commandId: 'pivotFieldSettings',
          ...(numberFormatPicker ? { pickNumberFormat: numberFormatPicker.pick } : {}),
          onChanged: () => {
            deps.onAfterPivotEdit?.();
            refresh();
          },
        },
        addr,
      ),
    openPivotOptions: (sheetIndex, pivotIndex) => {
      if (wb.pivotTableName(sheetIndex, pivotIndex) === null) return false;
      void openPivotOptions(
        {
          host,
          strings,
          wb,
          history: deps.history ?? null,
          store: deps.store,
          origin: 'ribbon',
          commandId: 'pivotOptions',
          onChanged: () => {
            deps.onAfterPivotEdit?.();
            refresh();
          },
        },
        sheetIndex,
        pivotIndex,
      );
      return true;
    },
    isPivotFieldListOpen: () => open && activePivotFieldListKey.length > 0,
    hidePivotFieldList: () => {
      pivotFieldListHidden = true;
      close();
    },
    isPivotFieldListHidden: () => pivotFieldListHidden,
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
      unsubscribeSessionObjects?.();
      root.removeEventListener('keydown', onKey);
      detachPaneEscape();
      root.hidden = true;
      placeRoot(false);
      numberFormatPicker?.detach();
      root.remove();
    },
  };
}
