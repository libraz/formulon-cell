import { summarizeSpreadsheetCompatibility } from '../engine/compatibility.js';
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
import type { SessionIllustration } from '../store/store.js';
import { appendDialogIconButton } from './dialog-shell.js';
import { compatibilityLabelKey } from './spreadsheet-compatibility-report.js';
import { createWorkbookObjectsActionButton, pivotEditField } from './workbook-objects-dom.js';
import { type PivotEditorContext, renderPivotEditForm } from './workbook-objects-pivot-editor.js';

export {
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
  isPivotFieldListOpen(): boolean;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  bindWorkbook(next: WorkbookHandle): void;
  detach(): void;
}

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
  let pivotEditError = '';

  const root = document.createElement('div');
  root.className = 'fc-objects';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'false');
  root.hidden = true;
  root.tabIndex = -1;
  host.appendChild(root);

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
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions fc-objects__illustration-actions';
    const select = createWorkbookObjectsActionButton(strings.workbookObjects.objectSelect);
    select.addEventListener('click', () => deps.onSelectSessionIllustration?.(illustration.id));
    actions.appendChild(select);
    if (deps.onDuplicateSessionIllustration) {
      const duplicate = createWorkbookObjectsActionButton(strings.workbookObjects.objectDuplicate);
      duplicate.addEventListener('click', () =>
        deps.onDuplicateSessionIllustration?.(illustration.id),
      );
      actions.appendChild(duplicate);
    }
    if (deps.onClearSessionIllustration) {
      const remove = createWorkbookObjectsActionButton(strings.workbookObjects.objectDelete);
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
    color.value = illustration.color ?? '#0f6cbd';
    const radius = document.createElement('input');
    radius.className = 'fc-objects__input';
    radius.type = 'number';
    radius.min = '0';
    radius.max = '48';
    radius.step = '1';
    radius.value = String(
      illustration.radius ?? (illustration.shape === 'rounded-rectangle' ? 12 : 0),
    );
    const lineWidth = document.createElement('input');
    lineWidth.className = 'fc-objects__input';
    lineWidth.type = 'number';
    lineWidth.min = '1';
    lineWidth.max = '16';
    lineWidth.step = '1';
    lineWidth.value = String(illustration.lineWidth ?? 3);
    const opacity = document.createElement('input');
    opacity.className = 'fc-objects__input';
    opacity.type = 'range';
    opacity.min = '0';
    opacity.max = '1';
    opacity.step = '0.05';
    opacity.value = String(illustration.opacity ?? 0.16);
    const actions = document.createElement('div');
    actions.className = 'fc-objects__actions';
    const apply = createWorkbookObjectsActionButton(t.apply, { primary: true, type: 'submit' });
    actions.appendChild(apply);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      deps.onUpdateSessionIllustration?.(illustration.id, {
        color: color.value,
        radius: Math.max(0, Math.min(48, Number(radius.value) || 0)),
        lineWidth: Math.max(1, Math.min(16, Number(lineWidth.value) || 1)),
        opacity: Math.max(0, Math.min(1, Number(opacity.value) || 0)),
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
      deps.onAfterPivotEdit?.();
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
    root.replaceChildren();
    root.className = `fc-objects${activeFieldListPivot ? ' fc-objects--taskpane' : ''}`;
    root.setAttribute('aria-label', activeFieldListPivot ? t.pivotFieldList : t.title);

    const header = document.createElement('div');
    header.className = 'fc-objects__header';
    const title = document.createElement('div');
    title.className = 'fc-objects__title';
    title.textContent = activeFieldListPivot ? t.pivotFieldList : t.title;
    const headerActions = document.createElement('div');
    headerActions.className = 'fc-objects__header-actions';
    if (activeFieldListPivot) {
      const back = createWorkbookObjectsActionButton(t.backToWorkbookObjects);
      back.addEventListener('click', () => {
        activePivotFieldListKey = '';
        pivotEditError = '';
        render();
      });
      headerActions.appendChild(back);
    }
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
      const section = document.createElement('section');
      section.className = 'fc-objects__section';
      const heading = document.createElement('div');
      heading.className = 'fc-objects__heading';
      heading.textContent = [
        `${t.pivot} ${activeFieldListPivot.pivotIndex + 1}`,
        `${t.sheet} ${activeFieldListPivot.sheetIndex + 1}`,
        pivotAnchor(activeFieldListPivot),
      ].join(' · ');
      section.append(
        heading,
        renderPivotEditForm(pivotEditorContext(), activeFieldListPivot, { fieldListOnly: true }),
      );
      body.appendChild(section);
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
        if (
          canEditPivot &&
          activePivotFieldListKey === pivotKey(pivot.sheetIndex, pivot.pivotIndex)
        ) {
          card.appendChild(
            renderPivotEditForm(pivotEditorContext(), pivot, { fieldListOnly: true }),
          );
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
      for (const [index, illustration] of illustrations.entries()) {
        const li = document.createElement('li');
        li.textContent = [
          illustration.kind === 'image'
            ? strings.ribbon.pictures
            : (illustration.shape ?? strings.ribbon.shapes),
          `${t.sheet} ${illustration.sheet + 1}`,
          illustration.id || `${strings.ribbon.illustrations} ${index + 1}`,
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

  const refresh = (): void => {
    if (open) render();
  };
  const unsubscribeSessionObjects = deps.subscribeSessionObjects?.(refresh) ?? null;

  const openPanel = (): void => {
    render();
    restoreFocusEl = document.activeElement instanceof HTMLElement ? document.activeElement : host;
    root.hidden = false;
    open = true;
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
    activePivotFieldListKey = key;
    pivotEditError = '';
    openPanel();
    return true;
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') close(true);
  };
  root.addEventListener('keydown', onKey);

  return {
    open: openPanel,
    openPivotFieldList,
    isPivotFieldListOpen: () => open && activePivotFieldListKey.length > 0,
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
      root.remove();
    },
  };
}
