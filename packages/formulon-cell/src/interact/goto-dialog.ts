import {
  findMatchingRanges,
  type GoToSpecialKind,
  type GoToSpecialOptions,
  selectionFromRanges,
} from '../commands/goto-special.js';
import { formatA1Range } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { rangeDefinedNames, resolveAreaText, splitAreas } from '../mount/ref-utils.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogButton,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';
import {
  clampNavigationAddr,
  clampNavigationRange,
  isNavigationAddrAllowed,
  scrollActiveIntoView,
  syncNavigationViewport,
} from './navigation-policy.js';
import { attachRangePickerButton } from './range-picker-control.js';
import { attachValidationAlert } from './validation.js';

export interface GoToDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so a Go To Special invocation always queries the live engine. */
  getWb: () => WorkbookHandle;
  strings?: Strings;
}

export interface GoToDialogHandle {
  open(mode?: 'go-to' | 'special'): void;
  close(): void;
  detach(): void;
}

/** Dialog options in the desktop two-column order. */
const LEFT_KINDS: readonly GoToSpecialKind[] = [
  'comments',
  'constants',
  'formulas',
  'blanks',
  'current-region',
  'current-array',
];
const RIGHT_KINDS: readonly GoToSpecialKind[] = [
  'row-differences',
  'column-differences',
  'precedents',
  'dependents',
  'last-cell',
  'visible-cells',
  'conditional-format',
  'data-validation',
];

const MAX_RECENT_REFERENCES = 4;

/** Areas last jumped to through the dialog, newest first, per mount. */
const recentReferences = new WeakMap<SpreadsheetStore, Range[][]>();

const quoteSheetName = (name: string): string =>
  /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

/**
 * Spreadsheet-style Go To / Go To Special dialog. Go To jumps to names,
 * recent references or typed (multi-area, sheet-qualified) references; its
 * Special button switches to the category picker, which selects every
 * matching cell as merged rectangles.
 *
 * Lifecycle mirrors the other modals: `attach…()` mounts a hidden overlay on
 * the host; `open()` shows it, `close()` hides it, `detach()` tears it down.
 * A Go To Special search with no hits (or too many) closes the dialog and
 * raises a modal alert, leaving the selection unchanged.
 */
export function attachGoToDialog(deps: GoToDialogDeps): GoToDialogHandle {
  const { host, store, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.goToDialog;
  let mode: 'go-to' | 'special' = 'special';

  const shell = createDialogShell({
    host,
    className: 'fc-goto',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  // Reuse the shared format-dialog skin for header/footer/btn styling.
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-goto__panel'],
  });
  const alertDialog = attachValidationAlert({
    host,
    labels: {
      ok: strings.formatDialog.ok,
      stop: strings.formatDialog.validationErrorStop,
      warning: strings.formatDialog.validationErrorWarning,
      information: strings.formatDialog.validationErrorInfo,
    },
  });

  // ── Go To: names and recent references ────────────────────────────────
  const listLegend = document.createElement('div');
  listLegend.className = 'fc-goto__legend';
  listLegend.textContent = t.goToLabel;
  body.appendChild(listLegend);

  const targetList = document.createElement('div');
  targetList.className = 'fc-goto__kinds fc-goto__targets';
  targetList.setAttribute('role', 'listbox');
  targetList.setAttribute('aria-label', t.goToLabel);
  body.appendChild(targetList);

  // ── Go To: direct reference ───────────────────────────────────────────
  const referenceRow = document.createElement('label');
  referenceRow.className = 'fc-goto__reference';
  const referenceLabel = document.createElement('span');
  referenceLabel.textContent = t.reference;
  const referenceInput = document.createElement('input');
  referenceInput.type = 'text';
  referenceInput.placeholder = t.referencePlaceholder;
  referenceInput.autocomplete = 'off';
  referenceInput.spellcheck = false;
  referenceRow.append(referenceLabel, referenceInput);
  attachRangePickerButton(referenceInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => formatA1Range(store.getState().selection.range),
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'go-to-reference',
  });
  body.appendChild(referenceRow);

  // ── Go To Special: category radios in two columns ─────────────────────
  const kindList = document.createElement('div');
  kindList.className = 'fc-goto__kinds';
  body.appendChild(kindList);
  const leftColumn = document.createElement('div');
  const rightColumn = document.createElement('div');
  kindList.append(leftColumn, rightColumn);

  const suffix = Math.random().toString(36).slice(2, 8);
  const kindName = `fc-goto-kind-${suffix}`;
  const kindLabels: Record<string, string> = {
    comments: t.kindComments,
    constants: t.kindConstants,
    formulas: t.kindFormulas,
    blanks: t.kindBlanks,
    'current-region': t.kindCurrentRegion,
    'current-array': t.kindCurrentArray,
    'row-differences': t.kindRowDifferences,
    'column-differences': t.kindColumnDifferences,
    precedents: t.kindPrecedents,
    dependents: t.kindDependents,
    'last-cell': t.kindLastCell,
    'visible-cells': t.kindVisibleCells,
    'conditional-format': t.kindConditionalFormat,
    'data-validation': t.kindDataValidation,
  };
  const makeRadio = (
    parent: HTMLElement,
    name: string,
    value: string,
    label: string,
    sub?: 'levels' | 'rule',
  ): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = name;
    input.value = value;
    if (sub) {
      input.dataset.sub = sub;
      wrap.style.paddingInlineStart = '20px';
    }
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    parent.appendChild(wrap);
    return input;
  };

  const kindInputs = new Map<GoToSpecialKind, HTMLInputElement>();
  const addKind = (parent: HTMLElement, kind: GoToSpecialKind): void => {
    kindInputs.set(kind, makeRadio(parent, kindName, kind, kindLabels[kind] as string));
  };

  // Excel shows the value-type checkboxes beneath Formulas; they narrow
  // Constants and Formulas alike.
  const valueFilters = document.createElement('div');
  valueFilters.className = 'fc-goto__value-filters';
  const makeValueFilter = (label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__check';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = true;
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    valueFilters.appendChild(wrap);
    return input;
  };
  const valueFilterInputs = {
    numbers: makeValueFilter(t.kindNumbers),
    text: makeValueFilter(t.kindText),
    logical: makeValueFilter(t.kindLogical),
    errors: makeValueFilter(t.kindErrors),
  };
  for (const [key, input] of Object.entries(valueFilterInputs)) input.value = key;

  for (const kind of LEFT_KINDS) {
    addKind(leftColumn, kind);
    if (kind === 'formulas') leftColumn.appendChild(valueFilters);
  }

  const levelName = `fc-goto-level-${suffix}`;
  const ruleName = `fc-goto-rule-${suffix}`;
  const levelInputs = {
    direct: makeRadio(
      document.createElement('div'),
      levelName,
      'direct',
      t.kindDirectOnly,
      'levels',
    ),
    all: makeRadio(document.createElement('div'), levelName, 'all', t.kindAllLevels, 'levels'),
  };
  const ruleInputs = {
    all: makeRadio(document.createElement('div'), ruleName, 'all', t.ruleAll, 'rule'),
    same: makeRadio(document.createElement('div'), ruleName, 'same', t.ruleSame, 'rule'),
  };
  for (const kind of RIGHT_KINDS) {
    addKind(rightColumn, kind);
    if (kind === 'dependents') {
      for (const input of Object.values(levelInputs))
        rightColumn.appendChild(input.parentElement as Node);
    }
    if (kind === 'data-validation') {
      for (const input of Object.values(ruleInputs))
        rightColumn.appendChild(input.parentElement as Node);
    }
  }
  // Defaults: constants, direct precedents/dependents, every rule cell.
  (kindInputs.get('constants') as HTMLInputElement).checked = true;
  levelInputs.direct.checked = true;
  ruleInputs.all.checked = true;

  // Inline status (shown for an invalid reference).
  const statusLine = document.createElement('div');
  statusLine.className = 'fc-goto__status';
  statusLine.setAttribute('role', 'status');
  statusLine.setAttribute('aria-live', 'polite');
  body.appendChild(statusLine);

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
  });
  const specialBtn = createDialogButton({ label: t.special, variant: 'secondary' });
  footer.insertBefore(specialBtn, footer.firstChild);

  const getCheckedKind = (): GoToSpecialKind => {
    for (const [k, input] of kindInputs) if (input.checked) return k;
    return 'constants';
  };
  const getOptions = (kind: GoToSpecialKind): GoToSpecialOptions => {
    if (kind === 'formulas' || kind === 'constants') {
      return {
        filters: {
          numbers: valueFilterInputs.numbers.checked,
          text: valueFilterInputs.text.checked,
          logical: valueFilterInputs.logical.checked,
          errors: valueFilterInputs.errors.checked,
        },
      };
    }
    if (kind === 'precedents' || kind === 'dependents')
      return { allLevels: levelInputs.all.checked };
    if (kind === 'conditional-format' || kind === 'data-validation') {
      return { same: ruleInputs.same.checked };
    }
    return {};
  };

  /** Resolve the typed text into areas on a single sheet; null when any part is invalid. */
  const resolveReference = (text: string): Range[] | null => {
    const wb = getWb();
    const current = store.getState().data.sheetIndex;
    const areas: Range[] = [];
    for (const part of splitAreas(text)) {
      const area = resolveAreaText(wb, part, current);
      if (!area) return null;
      areas.push(area);
    }
    const sheet = areas[0]?.sheet;
    if (sheet === undefined || areas.some((a) => a.sheet !== sheet)) return null;
    // A Go To reference is an explicit host/user target. Reject a partially
    // out-of-bounds rectangle instead of silently clamping it.
    const allowed = areas.every(
      (a) =>
        isNavigationAddrAllowed(store, { sheet, row: a.r0, col: a.c0 }) &&
        isNavigationAddrAllowed(store, { sheet, row: a.r1, col: a.c1 }),
    );
    if (!allowed) return null;
    const clamped = areas.map((a) => clampNavigationRange(store, a));
    return clamped.every((a): a is Range => a !== null) ? clamped : null;
  };

  /** `$A$1` on the active sheet, `Sheet2!$A$1` elsewhere; areas joined by commas. */
  const formatReference = (areas: readonly Range[]): string => {
    const current = store.getState().data.sheetIndex;
    const wb = getWb();
    return areas
      .map((a) => {
        const ref = formatA1Range(a, { absolute: true });
        return a.sheet === current ? ref : `${quoteSheetName(wb.sheetName(a.sheet))}!${ref}`;
      })
      .join(',');
  };

  const fillReference = (text: string): void => {
    referenceInput.value = text;
    statusLine.textContent = '';
    referenceInput.focus();
    referenceInput.select();
  };

  const renderTargets = (): void => {
    targetList.replaceChildren();
    const entries: string[] = [];
    for (const dn of rangeDefinedNames(getWb(), store.getState().data.sheetIndex)) {
      entries.push(dn.name);
    }
    for (const areas of recentReferences.get(store) ?? []) entries.push(formatReference(areas));
    for (const text of entries) {
      const item = createDialogButton({ label: text, baseClass: 'fc-goto__radio' });
      item.setAttribute('role', 'option');
      item.addEventListener('click', () => fillReference(text));
      item.addEventListener('dblclick', () => {
        fillReference(text);
        onOk();
      });
      targetList.appendChild(item);
    }
    targetList.hidden = entries.length === 0;
    listLegend.hidden = entries.length === 0;
  };

  const goToReference = (): boolean => {
    const areas = resolveReference(referenceInput.value);
    const [first, ...rest] = areas ?? [];
    const target = first
      ? clampNavigationAddr(store, { sheet: first.sheet, row: first.r0, col: first.c0 })
      : null;
    if (!areas || !first || !target) {
      statusLine.textContent = t.invalidReference;
      return false;
    }
    store.setState((s) => ({
      ...s,
      data: { ...s.data, sheetIndex: first.sheet },
      selection: {
        active: target,
        anchor: target,
        range: first,
        extraRanges: rest,
      },
    }));
    syncNavigationViewport(store);
    scrollActiveIntoView(store, target);
    const text = formatReference(areas);
    const recents = (recentReferences.get(store) ?? []).filter(
      (prev) => formatReference(prev) !== text,
    );
    recentReferences.set(store, [areas, ...recents].slice(0, MAX_RECENT_REFERENCES));
    api.close();
    return true;
  };

  const goToSpecial = (): void => {
    const kind = getCheckedKind();
    const wb = getWb();
    const result = findMatchingRanges(wb, store, kind, getOptions(kind));
    const ranges =
      result.status === 'found'
        ? result.ranges
            .map((r) => clampNavigationRange(store, r))
            .filter((r): r is Range => r !== null)
        : [];
    api.close();
    if (result.status === 'too-large') {
      alertDialog.show({ severity: 'stop', title: t.title, message: t.selectionTooLarge });
      return;
    }
    if (ranges.length === 0) {
      alertDialog.show({ severity: 'stop', title: t.title, message: t.noResults });
      return;
    }
    store.setState((s) => ({ ...s, selection: selectionFromRanges(ranges) }));
  };

  const onOk = (): void => {
    statusLine.textContent = '';
    if (mode === 'go-to') goToReference();
    else goToSpecial();
  };

  const syncEnabled = (): void => {
    const kind = getCheckedKind();
    const filtered = kind === 'formulas' || kind === 'constants';
    for (const input of Object.values(valueFilterInputs))
      projectDisabledState(input, !filtered, null);
    const traced = kind === 'precedents' || kind === 'dependents';
    for (const input of Object.values(levelInputs)) projectDisabledState(input, !traced, null);
    const ruled = kind === 'conditional-format' || kind === 'data-validation';
    for (const input of Object.values(ruleInputs)) projectDisabledState(input, !ruled, null);
  };

  for (const input of kindInputs.values()) shell.on(input, 'change', syncEnabled);

  const showMode = (next: 'go-to' | 'special'): void => {
    mode = next;
    statusLine.textContent = '';
    const normal = mode === 'go-to';
    header.textContent = normal ? t.goToTitle : t.title;
    shell.setAriaLabel(normal ? t.goToTitle : t.title);
    referenceRow.hidden = !normal;
    specialBtn.hidden = !normal;
    kindList.hidden = normal;
    if (normal) renderTargets();
    else {
      targetList.hidden = true;
      listLegend.hidden = true;
      syncEnabled();
    }
  };

  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', () => api.close());
  shell.on(specialBtn, 'click', () => {
    showMode('special');
    kindInputs.get(getCheckedKind())?.focus();
  });
  // Enter on any control inside the modal commits — matches the spreadsheet
  // convention. Escape and backdrop dismissals are handled by the shell.
  shell.on(shell.overlay, 'keydown', (e) => {
    if ((e as KeyboardEvent).key !== 'Enter') return;
    // Buttons keep their own click behavior.
    if ((e.target as HTMLElement).closest('button')) return;
    e.preventDefault();
    onOk();
  });

  const api: GoToDialogHandle = {
    open(nextMode = 'special'): void {
      showMode(nextMode);
      if (nextMode === 'go-to') referenceInput.value = '';
      shell.open();
      deferInitialFocus(shell.overlay, () => {
        if (nextMode === 'go-to') {
          referenceInput.focus();
          return;
        }
        kindInputs.get(getCheckedKind())?.focus();
      });
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      alertDialog.detach();
      shell.dispose();
    },
  };
  return api;
}
