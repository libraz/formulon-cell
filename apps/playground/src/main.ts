import {
  aggregateSelection,
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyMerge,
  applyTextScript,
  applyUnmerge,
  attachFilterDropdown,
  autoSum,
  buildRibbonModel,
  bumpDecimals,
  type CellBorderStyle,
  clearFormat,
  copy,
  createSessionChart,
  cut,
  cycleCurrency,
  cyclePercent,
  deleteCols,
  deleteRows,
  fluentIconPaths,
  formatAsTable,
  hideCols,
  hideRows,
  insertCols,
  insertRows,
  mutators,
  type NumFmt,
  parseScriptCommand,
  pasteTSV,
  projectActiveState,
  type ReviewCell,
  RIBBON_KEYSHORTCUTS,
  type RibbonCommand,
  type RibbonReportItem,
  type RibbonTab,
  recordFormatChange,
  removeDuplicates,
  Spreadsheet,
  type SpreadsheetInstance,
  setAlign,
  setBorderPreset,
  setFillColor,
  setFont,
  setFontColor,
  setFreezePanes,
  setNumFmt,
  setSheetHidden,
  setSheetZoom,
  setVAlign,
  sortRange,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
  toggleWrap,
  toolbarText,
  WorkbookHandle,
} from '@libraz/formulon-cell';
import { showMessage, showPrompt, showReport } from './dialogs.js';
import { applyFixture, isFixtureName } from './fixtures.js';
import { focusMenuItem, handleMenuKeydown, prepareMenu } from './menu-a11y.js';
import { seedWorkbook } from './seed.js';
import { openSheetTabMenu } from './sheet-tab-menu.js';
import { setupSortMenu, setupZoomControls } from './zoom-sort.js';

const sheetEl = document.getElementById('sheet');
const themeToggle = document.getElementById('theme-toggle') as HTMLButtonElement | null;
const themeLabel = document.getElementById('theme-label');
const docState = document.getElementById('doc-state');
const enginePill = document.getElementById('engine-pill');
const statusState = document.getElementById('status-state');
const statusSelection = document.getElementById('status-selection');
const statusMetric = document.getElementById('status-metric');
const statusEngine = document.getElementById('status-engine');
const statusObjects = document.getElementById('status-objects');
const ribbonRoot = document.getElementById('ribbon-root');

if (!sheetEl) throw new Error('#sheet missing');
if (statusMetric) {
  statusMetric.tabIndex = 0;
  statusMetric.setAttribute('aria-haspopup', 'menu');
}

// `paper` / `ink` are the core's theme names; the UI labels them Light / Dark.
type CoreTheme = 'paper' | 'ink';
type UiTheme = 'light' | 'dark';

const html = document.documentElement;
// URL params: `?theme=light|dark` and `?locale=en|ja` let E2E / visual specs
// pin the boot state without scripting the toolbar. They simply override the
// initial values; user toggles still work afterwards.
const bootParams = new URLSearchParams(window.location.search);
const themeParam = bootParams.get('theme');
const localeParam = bootParams.get('locale');
const initialUiTheme: UiTheme =
  themeParam === 'dark' || themeParam === 'light'
    ? themeParam
    : ((html.dataset.theme as UiTheme | undefined) ?? 'light');
let uiTheme: UiTheme = initialUiTheme;
html.dataset.theme = uiTheme;
const toCore = (t: UiTheme): CoreTheme => (t === 'dark' ? 'ink' : 'paper');

let inst: SpreadsheetInstance | null = null;

const seed = seedWorkbook;

const ribbonLang = localeParam === 'ja' ? 'ja' : 'en';
const ribbonText = toolbarText(ribbonLang);
let activeRibbonTab: RibbonTab = 'home';
let selectedBorderStyle: CellBorderStyle = 'thin';
let selectedBorderColor = '#000000';

const legacyCommandIds: Record<string, string> = {
  alignC: 'btn-align-center',
  alignL: 'btn-align-left',
  alignR: 'btn-align-right',
  autosum: 'btn-autosum',
  bold: 'btn-bold',
  borders: 'btn-borders',
  currency: 'btn-currency',
  decDown: 'btn-decimals-down',
  decUp: 'btn-decimals-up',
  filter: 'btn-sort',
  fontGrow: 'btn-font-grow',
  fontShrink: 'btn-font-shrink',
  formatPainter: 'btn-format-painter',
  freeze: 'btn-freeze',
  italic: 'btn-italic',
  merge: 'btn-merge',
  middle: 'btn-middle',
  percent: 'btn-percent',
  comma: 'btn-comma',
  commentInsert: 'btn-comment',
  hyperlinkInsert: 'btn-hyperlink',
  newCommentReview: 'btn-review-comment',
  pivotTableInsert: 'btn-pivot',
  redoHome: 'btn-redo',
  strike: 'btn-strike',
  top: 'btn-top',
  underline: 'btn-underline',
  undoHome: 'btn-undo',
  wrap: 'btn-wrap',
};

const renderRibbon = (): void => {
  if (!ribbonRoot) return;
  const model = buildRibbonModel(ribbonLang);
  const shell = document.createElement('div');
  shell.className = 'demo__ribbon-shell app__ribbon-shell';

  const tabs = document.createElement('div');
  tabs.className = 'demo__ribbon-tabs';
  tabs.setAttribute('role', 'tablist');
  tabs.setAttribute('aria-label', ribbonText.ribbonTabs);
  for (const tab of model) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `demo__ribbon-tab${tab.id === 'file' ? ' demo__ribbon-tab--file' : ''}${
      tab.id === activeRibbonTab ? ' demo__ribbon-tab--active' : ''
    }`;
    btn.setAttribute('role', 'tab');
    btn.setAttribute('aria-selected', tab.id === activeRibbonTab ? 'true' : 'false');
    btn.tabIndex = tab.id === activeRibbonTab ? 0 : -1;
    btn.dataset.ribbonTab = tab.id;
    btn.textContent = tab.label;
    tabs.appendChild(btn);
  }
  shell.appendChild(tabs);

  for (const tab of model) {
    const panel = document.createElement('div');
    panel.className = 'demo__ribbon';
    panel.setAttribute('role', 'toolbar');
    panel.setAttribute('aria-label', `${tab.label} ${ribbonText.ribbon}`);
    panel.dataset.ribbonPanel = tab.id;
    panel.hidden = tab.id !== activeRibbonTab;

    for (const g of tab.groups) {
      const group = document.createElement('section');
      group.className = `demo__ribbon-group${g.variant ? ` demo__ribbon-group--${g.variant}` : ''}`;
      group.setAttribute('aria-label', g.title);

      const tools = document.createElement('div');
      tools.className = 'demo__ribbon-tools';
      for (const c of g.commands) {
        if (c.kind === 'break') {
          const rowBreak = document.createElement('div');
          rowBreak.className = 'demo__rb-break';
          rowBreak.dataset.ribbonCommand = c.id;
          tools.appendChild(rowBreak);
          continue;
        }
        if (c.kind === 'select') {
          tools.appendChild(createRibbonSelect(c));
          continue;
        }
        if (c.kind === 'color') {
          tools.appendChild(createRibbonColor(c));
          continue;
        }
        const b = document.createElement('button');
        b.type = 'button';
        b.className = `demo__rb${c.kind === 'large' ? ' demo__rb--large' : ''}${
          c.kind === 'wide' ? ' demo__rb--wide' : ''
        }${c.kind === 'mono' ? ' demo__rb--mono' : ''}`;
        b.title = c.title;
        b.setAttribute('aria-label', c.title);
        const keyshortcuts = RIBBON_KEYSHORTCUTS[c.id];
        if (keyshortcuts) b.setAttribute('aria-keyshortcuts', keyshortcuts);
        b.dataset.ribbonCommand = c.id;
        const legacyId = legacyCommandIds[c.id];
        if (legacyId) b.id = legacyId;
        b.disabled = !!c.disabled;
        const textOnly = !c.icon || c.kind === 'mono';
        const showLabel = textOnly || c.kind === 'wide' || c.kind === 'large';
        const icon = c.icon && c.kind !== 'mono' ? createRibbonIcon(c.icon) : null;
        if (icon) {
          b.appendChild(icon);
        }
        if (showLabel || (!icon && c.kind !== 'mono')) {
          const label = document.createElement('span');
          label.textContent = c.label;
          b.appendChild(label);
        }
        tools.appendChild(b);
        if (c.id === 'borders') tools.appendChild(createBordersMenu());
        else if (c.id === 'freeze') tools.appendChild(createFreezeMenu());
        else if (c.id === 'filter') tools.appendChild(createSortMenu());
      }

      const label = document.createElement('div');
      label.className = 'demo__ribbon-label';
      label.textContent = g.title;
      group.appendChild(tools);
      group.appendChild(label);
      panel.appendChild(group);
    }

    shell.appendChild(panel);
  }

  ribbonRoot.replaceChildren(shell);
};

const createRibbonIcon = (name: string): SVGSVGElement | null => {
  const paths = fluentIconPaths(name);
  if (!paths) return null;
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('demo__rb-icon');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('fill', 'currentColor');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  for (const d of paths) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  return svg;
};

const activeCellFormat = () => {
  if (!inst) return null;
  const s = inst.store.getState();
  const a = s.selection.active;
  return s.format.formats.get(`${a.sheet}:${a.row}:${a.col}`) ?? null;
};

const currentRibbonControlValue = (id: string): string => {
  const f = activeCellFormat();
  if (id === 'fontFamily') return f?.fontFamily ?? 'Aptos';
  if (id === 'fontSize') return String(f?.fontSize ?? 11);
  if (id === 'fontColor') return f?.color ?? '#201f1e';
  if (id === 'fillColor') return f?.fill ?? '#ffffff';
  if (id === 'numberFormat') return inst ? projectActiveState(inst).numberFormat : 'general';
  if (id === 'marginsPreset') return 'normal';
  if (id === 'orientationPreset') return 'portrait';
  if (id === 'paperSizePreset') return 'A4';
  return '';
};

const numberFormatForAction = (action: string): NumFmt | null => {
  const symbol = ribbonLang === 'ja' ? '¥' : '$';
  switch (action) {
    case 'general':
      return { kind: 'general' };
    case 'fixed':
      return { kind: 'fixed', decimals: 0 };
    case 'currency':
      return { kind: 'currency', decimals: 0, symbol };
    case 'accounting':
      return { kind: 'accounting', decimals: 0, symbol };
    case 'shortDate':
      return { kind: 'date', pattern: ribbonLang === 'ja' ? 'yyyy/m/d' : 'm/d/yyyy' };
    case 'longDate':
      return { kind: 'date', pattern: ribbonLang === 'ja' ? 'yyyy"年"m"月"d"日' : 'mmmm d, yyyy' };
    case 'time':
      return { kind: 'time', pattern: ribbonLang === 'ja' ? 'H:MM' : 'h:MM AM/PM' };
    case 'percent':
      return { kind: 'percent', decimals: 0 };
    case 'fraction':
      return { kind: 'custom', pattern: '# ?/?' };
    case 'scientific':
      return { kind: 'scientific', decimals: 2 };
    case 'text':
      return { kind: 'text' };
    case 'more':
      return null;
    default:
      return { kind: 'general' };
  }
};

function applyRibbonFormat(
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void {
  const i = inst;
  if (!i) return;
  recordFormatChange(i.history, i.store, () => {
    fn(i.store.getState(), i.store);
  });
  (sheetEl as HTMLElement).focus();
}

function applyRibbonControl(id: string, value: string): void {
  if (id === 'fontFamily') {
    applyRibbonFormat((state, store) => setFont(state, store, { fontFamily: value }));
  } else if (id === 'fontSize') {
    applyRibbonFormat((state, store) => setFont(state, store, { fontSize: Number(value) }));
  } else if (id === 'fontColor') {
    applyRibbonFormat((state, store) => setFontColor(state, store, value));
  } else if (id === 'fillColor') {
    applyRibbonFormat((state, store) => setFillColor(state, store, value));
  } else if (id === 'numberFormat') {
    if (value === 'more') {
      inst?.openFormatDialog();
      return;
    }
    const fmt = numberFormatForAction(value);
    if (fmt) applyRibbonFormat((state, store) => setNumFmt(state, store, fmt));
  } else if (id === 'marginsPreset' || id === 'orientationPreset' || id === 'paperSizePreset') {
    inst?.openPageSetup();
  }
}

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
  const ja = ribbonLang === 'ja';
  const fmt = (top: string, bottom: string, left: string, right: string): string =>
    ja
      ? `上: ${top}", 下: ${bottom}", 左: ${left}", 右: ${right}"`
      : `Top: ${top}", Bottom: ${bottom}", Left: ${left}", Right: ${right}"`;
  switch (value) {
    case 'normal':
      return fmt('0.75', '0.75', '0.7', '0.7');
    case 'wide':
      return fmt('1', '1', '1', '1');
    case 'narrow':
      return fmt('0.75', '0.75', '0.25', '0.25');
    case 'custom':
      return ja ? 'ユーザー設定の余白...' : 'Custom margins...';
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

const themeFontValues = new Set(['Aptos', 'Aptos Display', 'Aptos Narrow']);
const recentFontValues = new Set(['Yu Gothic UI']);
const commonFontValues = new Set(['Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Consolas']);
const fontSubmenuFamilies = new Set(['Yu Gothic UI', 'BIZ UDGothic', 'Meiryo UI']);
const fontAvailabilityCache = new Map<string, boolean>();

const isJapaneseFontName = (value: string): boolean => /[\u3000-\u9fff]/.test(value);

const fontProbeContext = (): CanvasRenderingContext2D | null => {
  const canvas = document.createElement('canvas');
  return canvas.getContext('2d');
};

const isFontProbablyAvailable = (font: string): boolean => {
  if (themeFontValues.has(font) || commonFontValues.has(font)) return true;
  const cached = fontAvailabilityCache.get(font);
  if (cached !== undefined) return cached;
  const ctx = fontProbeContext();
  if (!ctx) return true;
  const sample = 'mmmmmmmmmwwwwwiiiiii 0123456789 あいう漢字';
  const available = ['serif', 'sans-serif', 'monospace'].some((fallback) => {
    ctx.font = `16px ${fallback}`;
    const fallbackWidth = ctx.measureText(sample).width;
    ctx.font = `16px "${font}", ${fallback}`;
    return Math.abs(ctx.measureText(sample).width - fallbackWidth) > 0.5;
  });
  fontAvailabilityCache.set(font, available);
  return available;
};

const shouldShowFontOption = (value: string, current: string): boolean => {
  if (value === current) return true;
  if (ribbonLang !== 'ja' && isJapaneseFontName(value)) return false;
  return isFontProbablyAvailable(value);
};

const ribbonFontSection = (
  value: string,
  options: readonly { value: string; label: string }[],
): string | null => {
  const firstTheme = options.find((option) => themeFontValues.has(option.value))?.value;
  if (value === firstTheme) return ribbonLang === 'ja' ? 'テーマのフォント' : 'Theme Fonts';
  const firstRecent = options.find((option) => recentFontValues.has(option.value))?.value;
  if (value === firstRecent)
    return ribbonLang === 'ja' ? '最近使ったフォント' : 'Recently Used Fonts';
  const firstAll = options.find(
    (option) => !themeFontValues.has(option.value) && !recentFontValues.has(option.value),
  )?.value;
  if (value === firstAll) return ribbonLang === 'ja' ? 'すべてのフォント' : 'All Fonts';
  return null;
};

const ribbonFontRole = (value: string): string | null => {
  switch (value) {
    case 'Aptos Display':
    case '游ゴシック Light':
      return ribbonLang === 'ja' ? '(見出し)' : '(Heading)';
    case 'Aptos Narrow':
    case '游ゴシック Regular':
      return ribbonLang === 'ja' ? '(本文)' : '(Body)';
    default:
      return null;
  }
};

const ribbonOptionsForCommand = (
  command: RibbonCommand,
  current: string,
): readonly { value: string; label: string }[] => {
  const options = command.options ?? [];
  if (command.id !== 'fontFamily') return options;
  return options.filter((option) => shouldShowFontOption(option.value, current));
};

const createRibbonSelect = (command: RibbonCommand): HTMLDivElement => {
  const wrap = document.createElement('div');
  wrap.className = `demo__rb-dd${command.className ? ` ${command.className}` : ''}`;
  wrap.dataset.ribbonCommand = command.id;
  wrap.dataset.ribbonSelect = command.id;
  wrap.dataset.ribbonOptions = JSON.stringify(command.options ?? []);

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'demo__rb-dd__btn';
  button.title = command.title;
  button.setAttribute('aria-label', command.title);
  button.setAttribute('aria-haspopup', 'listbox');
  button.setAttribute('aria-expanded', 'false');

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
      const section = command.id === 'fontFamily' ? ribbonFontSection(option.value, options) : null;
      if (section) {
        const heading = document.createElement('div');
        heading.className = 'demo__rb-dd__section';
        heading.setAttribute('role', 'presentation');
        heading.textContent = section;
        list.appendChild(heading);
      }
      const selected = option.value === current;
      const item = document.createElement('button');
      item.type = 'button';
      item.className = `demo__rb-dd__opt${selected ? ' demo__rb-dd__opt--selected' : ''}`;
      item.setAttribute('role', 'option');
      item.setAttribute('aria-selected', selected ? 'true' : 'false');
      item.tabIndex = -1;
      item.dataset.value = option.value;
      item.dataset.fcValue = option.value;
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
        if (fontSubmenuFamilies.has(option.value)) {
          const arrow = document.createElement('span');
          arrow.className = 'demo__rb-dd__submenu';
          arrow.setAttribute('aria-hidden', 'true');
          arrow.textContent = '›';
          item.appendChild(arrow);
        }
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

const createRibbonColor = (command: RibbonCommand): HTMLLabelElement => {
  const label = document.createElement('label');
  label.className = 'demo__rb-color';
  label.title = command.title;
  label.dataset.ribbonCommand = command.id;
  if (command.icon) {
    const icon = createRibbonIcon(command.icon);
    if (icon) label.appendChild(icon);
  }
  const input = document.createElement('input');
  input.type = 'color';
  input.setAttribute('aria-label', command.title);
  input.value = currentRibbonControlValue(command.id);
  input.addEventListener('change', () => applyRibbonControl(command.id, input.value));
  label.appendChild(input);
  return label;
};

const closeOpenRibbonDropdowns = (except?: HTMLElement): void => {
  for (const open of document.querySelectorAll<HTMLElement>('.demo__rb-dd--open')) {
    if (except && open === except) continue;
    open.classList.remove('demo__rb-dd--open');
    open
      .querySelector<HTMLButtonElement>('.demo__rb-dd__btn')
      ?.setAttribute('aria-expanded', 'false');
    open.querySelector('.demo__rb-dd__list')?.remove();
  }
};

const updateRibbonSelectDisplay = (wrap: HTMLElement, command: RibbonCommand): void => {
  const current = currentRibbonControlValue(command.id);
  const option = command.options?.find((candidate) => candidate.value === current);
  const value = wrap.querySelector<HTMLElement>('.demo__rb-dd__value');
  if (value) value.textContent = option?.label ?? current;
};

const ribbonSelectLabel = (wrap: HTMLElement, current: string): string => {
  try {
    const options = JSON.parse(wrap.dataset.ribbonOptions ?? '[]') as {
      value: string;
      label: string;
    }[];
    return options.find((option) => option.value === current)?.label ?? current;
  } catch {
    return current;
  }
};

const createMenu = (id: string): HTMLDivElement => {
  const menu = document.createElement('div');
  menu.className = 'app__menu';
  menu.id = id;
  menu.hidden = true;
  prepareMenu(menu);
  return menu;
};

const menuButton = (label: string, attr: string, value: string): HTMLButtonElement => {
  const button = document.createElement('button');
  button.className = 'app__menu-item';
  button.type = 'button';
  button.setAttribute('role', 'menuitem');
  button.dataset[attr] = value;
  button.textContent = label;
  return button;
};

// ── Borders dropdown (Excel-365 parity) ─────────────────────────────────
// Renders a small SVG cell-outline icon for each border preset. Sides are
// drawn solid in the foreground color (thin/thick/double); the unset sides
// show as a faint dashed cell-outline base so the user can still see the
// cell shape.
type BorderPreviewSide = 'thin' | 'thick' | 'double' | null;
interface BorderPreviewSpec {
  top?: BorderPreviewSide;
  right?: BorderPreviewSide;
  bottom?: BorderPreviewSide;
  left?: BorderPreviewSide;
  /** When true draws an inner cross dividing the icon into 2×2 cells
   *  (used for the "格子 / All borders" icon). */
  innerGrid?: boolean;
  /** When false suppresses the faint dashed cell-outline base. */
  showBase?: boolean;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const BORDER_ICON_BOX = { x: 2, y: 2, w: 12, h: 12 } as const;
const BORDER_BASE_COLOR = '#c7c7c7';

const drawBorderEdge = (
  svg: SVGSVGElement,
  side: BorderPreviewSide,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): void => {
  if (!side) return;
  const isHorizontal = y1 === y2;
  const mk = (offset = 0, width = 1): void => {
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', String(isHorizontal ? x1 : x1 + offset));
    line.setAttribute('y1', String(isHorizontal ? y1 + offset : y1));
    line.setAttribute('x2', String(isHorizontal ? x2 : x2 + offset));
    line.setAttribute('y2', String(isHorizontal ? y2 + offset : y2));
    line.setAttribute('stroke', 'currentColor');
    line.setAttribute('stroke-width', String(width));
    line.setAttribute('stroke-linecap', 'square');
    svg.appendChild(line);
  };
  if (side === 'thin') mk(0, 1);
  else if (side === 'thick') mk(0, 2);
  else if (side === 'double') {
    mk(-1, 0.75);
    mk(1, 0.75);
  }
};

const createBorderPreview = (spec: BorderPreviewSpec): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('app__border-preview');
  const { x, y, w, h } = BORDER_ICON_BOX;
  if (spec.showBase !== false) {
    const base = document.createElementNS(SVG_NS, 'rect');
    base.setAttribute('x', String(x + 0.5));
    base.setAttribute('y', String(y + 0.5));
    base.setAttribute('width', String(w - 1));
    base.setAttribute('height', String(h - 1));
    base.setAttribute('fill', 'none');
    base.setAttribute('stroke', BORDER_BASE_COLOR);
    base.setAttribute('stroke-width', '1');
    base.setAttribute('stroke-dasharray', '1.5 1.5');
    svg.appendChild(base);
  }
  if (spec.innerGrid) {
    const v = document.createElementNS(SVG_NS, 'line');
    v.setAttribute('x1', String(x + w / 2));
    v.setAttribute('y1', String(y));
    v.setAttribute('x2', String(x + w / 2));
    v.setAttribute('y2', String(y + h));
    v.setAttribute('stroke', 'currentColor');
    v.setAttribute('stroke-width', '1');
    svg.appendChild(v);
    const hLine = document.createElementNS(SVG_NS, 'line');
    hLine.setAttribute('x1', String(x));
    hLine.setAttribute('y1', String(y + h / 2));
    hLine.setAttribute('x2', String(x + w));
    hLine.setAttribute('y2', String(y + h / 2));
    hLine.setAttribute('stroke', 'currentColor');
    hLine.setAttribute('stroke-width', '1');
    svg.appendChild(hLine);
  }
  drawBorderEdge(svg, spec.top ?? null, x, y, x + w, y);
  drawBorderEdge(svg, spec.bottom ?? null, x, y + h, x + w, y + h);
  drawBorderEdge(svg, spec.left ?? null, x, y, x, y + h);
  drawBorderEdge(svg, spec.right ?? null, x + w, y, x + w, y + h);
  return svg;
};

const BORDER_PRESETS: Record<string, BorderPreviewSpec> = {
  bottom: { bottom: 'thin' },
  top: { top: 'thin' },
  left: { left: 'thin' },
  right: { right: 'thin' },
  clear: {},
  all: {
    top: 'thin',
    right: 'thin',
    bottom: 'thin',
    left: 'thin',
    innerGrid: true,
    showBase: false,
  },
  outline: {
    top: 'thin',
    right: 'thin',
    bottom: 'thin',
    left: 'thin',
    showBase: false,
  },
  thickOutline: {
    top: 'thick',
    right: 'thick',
    bottom: 'thick',
    left: 'thick',
    showBase: false,
  },
  doubleBottom: { bottom: 'double' },
  thickBottom: { bottom: 'thick' },
  topAndBottom: { top: 'thin', bottom: 'thin' },
  topAndThickBottom: { top: 'thin', bottom: 'thick' },
  topAndDoubleBottom: { top: 'thin', bottom: 'double' },
};

const presetMenuItem = (presetKey: string, label: string): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.dataset.borderPreset = presetKey;
  const spec = BORDER_PRESETS[presetKey];
  if (spec) btn.appendChild(createBorderPreview(spec));
  else {
    const spacer = document.createElement('span');
    spacer.className = 'app__menu-item__icon-spacer';
    btn.appendChild(spacer);
  }
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const menuSeparator = (): HTMLDivElement => {
  const sep = document.createElement('div');
  sep.className = 'app__menu-sep';
  sep.setAttribute('role', 'separator');
  return sep;
};

const menuSectionHeader = (label: string): HTMLDivElement => {
  const el = document.createElement('div');
  el.className = 'app__menu-heading';
  el.setAttribute('role', 'presentation');
  el.textContent = label;
  return el;
};

const drawActionItem = (
  action: string,
  label: string,
  icon?: BorderPreviewSpec,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitemcheckbox');
  btn.setAttribute('aria-checked', 'false');
  btn.dataset.borderDraw = action;
  if (icon) btn.appendChild(createBorderPreview(icon));
  else {
    const spacer = document.createElement('span');
    spacer.className = 'app__menu-item__icon-spacer';
    btn.appendChild(spacer);
  }
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  return btn;
};

const submenuTrigger = (
  submenuKey: 'lineColor' | 'lineStyle',
  label: string,
): HTMLButtonElement => {
  const btn = document.createElement('button');
  btn.className = 'app__menu-item app__menu-item--preset app__menu-item--submenu';
  btn.type = 'button';
  btn.setAttribute('role', 'menuitem');
  btn.setAttribute('aria-haspopup', 'menu');
  btn.setAttribute('aria-expanded', 'false');
  btn.dataset.borderSubmenu = submenuKey;
  const spacer = document.createElement('span');
  spacer.className = 'app__menu-item__icon-spacer';
  btn.appendChild(spacer);
  const text = document.createElement('span');
  text.className = 'app__menu-item__text';
  text.textContent = label;
  btn.appendChild(text);
  const caret = document.createElement('span');
  caret.className = 'app__menu-item__caret';
  caret.setAttribute('aria-hidden', 'true');
  caret.textContent = '▶';
  btn.appendChild(caret);
  return btn;
};

// Visual line-style sample for the "線のスタイル" submenu (image #2).
const createLineSamplePreview = (style: CellBorderStyle | 'none'): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 80 12');
  svg.setAttribute('width', '80');
  svg.setAttribute('height', '12');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('app__line-sample');
  if (style === 'none') return svg;
  const draw = (yOffset: number, w: number, d: string | null): void => {
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', '4');
    line.setAttribute('y1', String(yOffset));
    line.setAttribute('x2', '76');
    line.setAttribute('y2', String(yOffset));
    line.setAttribute('stroke', 'currentColor');
    line.setAttribute('stroke-width', String(w));
    if (d) line.setAttribute('stroke-dasharray', d);
    svg.appendChild(line);
  };
  if (style === 'double') {
    draw(4, 1, null);
    draw(8, 1, null);
    return svg;
  }
  if (style === 'thin') draw(6, 1, null);
  else if (style === 'hair') draw(6, 0.5, null);
  else if (style === 'medium') draw(6, 1.75, null);
  else if (style === 'thick') draw(6, 2.5, null);
  else if (style === 'dotted') draw(6, 1, '1.2 1.6');
  else if (style === 'dashed') draw(6, 1, '3 2');
  else if (style === 'mediumDashed') draw(6, 1.75, '4 2');
  else if (style === 'dashDot') draw(6, 1, '4 2 1 2');
  else if (style === 'mediumDashDot') draw(6, 1.75, '4 2 1 2');
  else if (style === 'dashDotDot') draw(6, 1, '4 2 1 2 1 2');
  else if (style === 'mediumDashDotDot') draw(6, 1.75, '4 2 1 2 1 2');
  else if (style === 'slantDashDot') draw(6, 1, '4 2 2 2');
  return svg;
};

const LINE_STYLES_ALL: (CellBorderStyle | 'none')[] = [
  'none',
  'thin',
  'hair',
  'dotted',
  'dashed',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'medium',
  'double',
  'thick',
];

const createLineStyleSubmenu = (label: string): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = 'app__submenu app__submenu--line-style';
  submenu.id = 'menu-borders-line-style';
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  for (const value of LINE_STYLES_ALL) {
    const btn = document.createElement('button');
    btn.className = 'app__submenu-item';
    btn.type = 'button';
    btn.setAttribute('role', 'menuitemradio');
    btn.setAttribute('aria-checked', value === 'thin' ? 'true' : 'false');
    btn.dataset.borderLineStyle = value;
    if (value === 'none') {
      const span = document.createElement('span');
      span.textContent = ribbonText.lineStyleNone;
      span.className = 'app__submenu-item__text';
      btn.appendChild(span);
    } else {
      btn.appendChild(createLineSamplePreview(value));
    }
    submenu.appendChild(btn);
  }
  return submenu;
};

const LINE_COLOR_PALETTE = [
  '#000000',
  '#7f7f7f',
  '#c00000',
  '#ff0000',
  '#ffc000',
  '#ffff00',
  '#92d050',
  '#00b050',
  '#00b0f0',
  '#0070c0',
  '#002060',
  '#7030a0',
];

const createLineColorSubmenu = (label: string): HTMLDivElement => {
  const submenu = document.createElement('div');
  submenu.className = 'app__submenu app__submenu--line-color';
  submenu.id = 'menu-borders-line-color';
  submenu.setAttribute('role', 'menu');
  submenu.setAttribute('aria-label', label);
  submenu.hidden = true;
  const grid = document.createElement('div');
  grid.className = 'app__color-grid';
  for (const hex of LINE_COLOR_PALETTE) {
    const swatch = document.createElement('button');
    swatch.className = 'app__color-swatch';
    swatch.type = 'button';
    swatch.setAttribute('role', 'menuitemradio');
    swatch.setAttribute('aria-checked', hex === '#000000' ? 'true' : 'false');
    swatch.dataset.borderLineColor = hex;
    swatch.style.background = hex;
    swatch.setAttribute('aria-label', hex);
    grid.appendChild(swatch);
  }
  submenu.appendChild(grid);
  return submenu;
};

const createBordersMenu = (): HTMLDivElement => {
  const t = ribbonText;
  const menu = createMenu('menu-borders');
  menu.classList.add('app__menu--borders');
  menu.append(
    // Section 1: single-side edges (image 1: 下罫線 / 上罫線 / 左罫線 / 右罫線)
    presetMenuItem('bottom', t.bottomBorder),
    presetMenuItem('top', t.topBorder),
    presetMenuItem('left', t.leftBorder),
    presetMenuItem('right', t.rightBorder),
    menuSeparator(),
    // Section 2: frame presets
    presetMenuItem('clear', t.noBorder),
    presetMenuItem('all', t.allBorders),
    presetMenuItem('outline', t.outsideBorders),
    presetMenuItem('thickOutline', t.thickOutsideBorders),
    menuSeparator(),
    // Section 3: combined
    presetMenuItem('doubleBottom', t.doubleBottomBorder),
    presetMenuItem('thickBottom', t.thickBottomBorder),
    presetMenuItem('topAndBottom', t.topAndBottomBorder),
    presetMenuItem('topAndThickBottom', t.topAndThickBottomBorder),
    presetMenuItem('topAndDoubleBottom', t.topAndDoubleBottomBorder),
    // Section 4 header + draw actions
    menuSectionHeader(t.drawBordersHeading),
    drawActionItem('draw', t.drawBorder, { bottom: 'thin' }),
    drawActionItem('grid', t.drawBorderGrid, {
      top: 'thin',
      right: 'thin',
      bottom: 'thin',
      left: 'thin',
      innerGrid: true,
      showBase: false,
    }),
    drawActionItem('erase', t.eraseBorder),
    submenuTrigger('lineColor', t.lineColor),
    submenuTrigger('lineStyle', t.lineStyle),
    menuSeparator(),
    // Footer
    presetMenuItem('format', t.moreBorders),
  );
  // Submenus sit beside the main dropdown.
  menu.appendChild(createLineColorSubmenu(t.lineColor));
  menu.appendChild(createLineStyleSubmenu(t.lineStyle));
  return menu;
};

const createFreezeMenu = (): HTMLDivElement => {
  const menu = createMenu('menu-freeze');
  menu.append(
    menuButton('Freeze first row', 'freeze', 'row'),
    menuButton('Freeze first column', 'freeze', 'col'),
    menuButton('Freeze up to selection', 'freeze', 'selection'),
    menuButton('Unfreeze', 'freeze', 'off'),
  );
  return menu;
};

const createSortMenu = (): HTMLDivElement => {
  const menu = createMenu('menu-sort');
  menu.append(
    menuButton('Sort A → Z', 'sort', 'asc'),
    menuButton('Sort Z → A', 'sort', 'desc'),
    menuButton('Filter…', 'sort', 'filter'),
    menuButton('Clear filter', 'sort', 'filter-clear'),
    menuButton('Remove duplicates', 'sort', 'dedupe'),
    menuButton('Conditional formatting…', 'sort', 'conditional'),
    menuButton('Named ranges…', 'sort', 'named'),
  );
  return menu;
};

renderRibbon();

const colLabel = (n: number): string => {
  let out = '';
  let v = n;
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const fmt = (n: number): string => {
  if (!Number.isFinite(n)) return '—';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs < 0.01 || abs >= 1e9)) return n.toExponential(3);
  return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
};

type StatKey = 'sum' | 'avg' | 'count' | 'min' | 'max';
const STAT_KEYS: StatKey[] = ['sum', 'avg', 'count', 'min', 'max'];
const activeStats: Set<StatKey> = (() => {
  try {
    const saved = localStorage.getItem('fc-status-stats');
    if (saved) return new Set(JSON.parse(saved) as StatKey[]);
  } catch {}
  return new Set<StatKey>(['sum', 'avg', 'count']);
})();
const persistStats = (): void => {
  try {
    localStorage.setItem('fc-status-stats', JSON.stringify(Array.from(activeStats)));
  } catch {}
};

// Composite badge showing both passthrough OOXML parts and spreadsheet Tables.
// We accumulate the latest snapshot from each event and render together so
// switching workbooks doesn't leak stale numbers from the previous one.
const objectCounts = { passthroughs: 0, tables: 0, passByCat: {} as Record<string, number> };
function refreshObjectsBadge(
  source: 'passthroughs' | 'tables',
  detail: { count: number; byCategory?: Record<string, number> },
): void {
  if (source === 'passthroughs') {
    objectCounts.passthroughs = detail.count;
    objectCounts.passByCat = detail.byCategory ?? {};
  } else {
    objectCounts.tables = detail.count;
  }
  if (!statusObjects) return;
  const parts: string[] = [];
  if (objectCounts.tables > 0)
    parts.push(`${objectCounts.tables} table${objectCounts.tables === 1 ? '' : 's'}`);
  const charts = objectCounts.passByCat.charts ?? 0;
  const drawings = objectCounts.passByCat.drawings ?? 0;
  const pivots = objectCounts.passByCat.pivotTables ?? 0;
  if (charts > 0) parts.push(`${charts} chart${charts === 1 ? '' : 's'}`);
  if (drawings > 0) parts.push(`${drawings} drawing${drawings === 1 ? '' : 's'}`);
  if (pivots > 0) parts.push(`${pivots} pivot${pivots === 1 ? '' : 's'}`);
  if (parts.length === 0) {
    statusObjects.hidden = true;
    statusObjects.textContent = '';
    return;
  }
  statusObjects.hidden = false;
  statusObjects.textContent = `objects · ${parts.join(', ')}`;
  statusObjects.title = 'Read-only — loaded from .xlsx but not editable in formulon-cell';
}

function projectStatus(): void {
  if (!inst) return;
  const s = inst.store.getState();
  const a = s.selection.active;
  const r = s.selection.range;

  if (statusSelection) {
    if (r.r0 === r.r1 && r.c0 === r.c1) {
      statusSelection.textContent = `${colLabel(a.col)}${a.row + 1}`;
    } else {
      const tl = `${colLabel(r.c0)}${r.r0 + 1}`;
      const br = `${colLabel(r.c1)}${r.r1 + 1}`;
      const cells = (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);
      statusSelection.textContent = `${tl}:${br} · ${cells} cells`;
    }
  }

  if (statusMetric) {
    const stats = aggregateSelection(s);
    if (stats.numericCount === 0) {
      statusMetric.textContent = '';
    } else {
      const parts: string[] = [];
      if (activeStats.has('sum')) parts.push(`Sum ${fmt(stats.sum)}`);
      if (activeStats.has('avg')) parts.push(`Avg ${fmt(stats.avg)}`);
      if (activeStats.has('count')) parts.push(`Count ${stats.numericCount}`);
      if (activeStats.has('min')) parts.push(`Min ${fmt(stats.min)}`);
      if (activeStats.has('max')) parts.push(`Max ${fmt(stats.max)}`);
      statusMetric.textContent = parts.join(' · ');
    }
  }
}

// Right-click on the status metric → checkbox menu to toggle stats.
statusMetric?.addEventListener('contextmenu', (e) => {
  e.preventDefault();
  const opener =
    document.activeElement instanceof HTMLElement ? document.activeElement : statusMetric;
  const menu = document.createElement('div');
  menu.className = 'app__dropdown';
  prepareMenu(menu, 'Selection summary');
  menu.style.position = 'fixed';
  menu.style.left = `${e.clientX}px`;
  menu.style.bottom = `${window.innerHeight - e.clientY + 4}px`;
  menu.style.top = '';
  let cleanupMenuListeners = (): void => {};
  const closeMenu = (restoreFocus = false): void => {
    menu.remove();
    cleanupMenuListeners();
    if (restoreFocus) opener?.focus();
  };
  for (const key of STAT_KEYS) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'app__menu-item';
    item.setAttribute('role', 'menuitemcheckbox');
    item.setAttribute('aria-checked', activeStats.has(key) ? 'true' : 'false');
    item.tabIndex = -1;
    item.textContent = `${activeStats.has(key) ? '✓ ' : '  '}${key.toUpperCase()}`;
    item.addEventListener('click', () => {
      if (activeStats.has(key)) activeStats.delete(key);
      else activeStats.add(key);
      persistStats();
      projectStatus();
      const checked = activeStats.has(key);
      item.setAttribute('aria-checked', checked ? 'true' : 'false');
      item.textContent = `${checked ? '✓ ' : '  '}${key.toUpperCase()}`;
    });
    menu.appendChild(item);
  }
  const close = (ev: MouseEvent): void => {
    if (!menu.contains(ev.target as Node)) {
      closeMenu();
    }
  };
  menu.addEventListener('keydown', (event) => {
    handleMenuKeydown(event, menu, { close: closeMenu, restoreFocusTo: opener });
  });
  cleanupMenuListeners = () => document.removeEventListener('mousedown', close, true);
  document.body.appendChild(menu);
  focusMenuItem(menu);
  setTimeout(() => document.addEventListener('mousedown', close, true), 0);
});

const ACTIVE_CLASS = 'demo__rb--active';
const setActive = (id: string, on: boolean): void => {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.toggle(ACTIVE_CLASS, on);
};

function projectFormatToolbar(): void {
  if (!inst) return;
  const s = inst.store.getState();
  const a = s.selection.active;
  const key = `${a.sheet}:${a.row}:${a.col}`;
  const f = s.format.formats.get(key);
  setActive('btn-bold', !!f?.bold);
  setActive('btn-italic', !!f?.italic);
  setActive('btn-underline', !!f?.underline);
  setActive('btn-strike', !!f?.strike);
  setActive('btn-align-left', f?.align === 'left');
  setActive('btn-align-center', f?.align === 'center');
  setActive('btn-align-right', f?.align === 'right');
  setActive('btn-currency', f?.numFmt?.kind === 'currency');
  setActive('btn-percent', f?.numFmt?.kind === 'percent');
  for (const wrap of document.querySelectorAll<HTMLElement>('[data-ribbon-select]')) {
    const id = wrap.dataset.ribbonSelect;
    if (!id) continue;
    const current = currentRibbonControlValue(id);
    const value = wrap.querySelector<HTMLElement>('.demo__rb-dd__value');
    if (value) value.textContent = ribbonSelectLabel(wrap, current);
    for (const option of wrap.querySelectorAll<HTMLElement>('.demo__rb-dd__opt')) {
      const selected = option.dataset.value === current;
      option.classList.toggle('demo__rb-dd__opt--selected', selected);
      option.setAttribute('aria-selected', selected ? 'true' : 'false');
    }
  }
  const fontColorInput = document.querySelector<HTMLInputElement>(
    '[data-ribbon-command="fontColor"] input',
  );
  if (fontColorInput) fontColorInput.value = f?.color ?? '#201f1e';
  const fillColorInput = document.querySelector<HTMLInputElement>(
    '[data-ribbon-command="fillColor"] input',
  );
  if (fillColorInput) fillColorInput.value = f?.fill ?? '#ffffff';
}

async function boot(): Promise<void> {
  // Default to the real WASM engine. Pass ?engine=stub to force the JS stub
  // for explicit demos or behavior diffs.
  const params = new URLSearchParams(window.location.search);
  const preferStub = params.get('engine') === 'stub';
  const wb = await WorkbookHandle.createDefault({
    preferStub,
    onFallback: (reason) => {
      // eslint-disable-next-line no-console
      console.info('[formulon-cell]', reason);
    },
  });
  // mount.ts only runs `seed` on workbooks it owns. We construct `wb` here so
  // we can read `isStub` / `version` for the engine pill before mounting,
  // which means we have to seed the workbook ourselves. `?fixture=empty`
  // (used by E2E specs that need a deterministic blank workbook) skips this.
  if (bootParams.get('fixture') !== 'empty') {
    seed(wb);
  }

  inst = await Spreadsheet.mount(sheetEl as HTMLElement, {
    theme: toCore(uiTheme),
    workbook: wb,
    locale: localeParam === 'ja' ? 'ja' : 'en',
    features: {
      watchWindow: true,
    },
  });
  // Debug-only: expose for browser console / e2e poking. Safe to leave on the
  // playground build; the core package never references this global.
  (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst = inst;

  // Visual-regression fixtures. `?fixture=cf|sparkline|selection|frozen`
  // replaces the default seed with a deterministic shape.
  const fixtureParam = bootParams.get('fixture');
  if (fixtureParam && isFixtureName(fixtureParam)) {
    applyFixture(fixtureParam, wb, inst);
  }

  filterDropdown = attachFilterDropdown({ store: inst.store });

  // Read-only badge — chart/drawing/pivot counts and spreadsheet Tables. Hidden
  //  until the loaded workbook actually carries any of these objects.
  inst.host.addEventListener('fc:passthroughs', (ev) => {
    const e = ev as CustomEvent<{ count: number; byCategory: Record<string, number> }>;
    refreshObjectsBadge('passthroughs', e.detail);
  });
  inst.host.addEventListener('fc:tables', (ev) => {
    const e = ev as CustomEvent<{ count: number }>;
    refreshObjectsBadge('tables', e.detail);
  });
  // Header chevron click → mount.ts owns the `fc:openfilter` listener and
  // opens its own dropdown. The playground keeps its `filterDropdown` only
  // for the sort menu's "filter" action.

  const engineLabel = wb.isStub ? 'stub engine' : `formulon ${wb.version}`;
  if (enginePill) enginePill.textContent = `engine · ${engineLabel}`;
  if (statusEngine) statusEngine.textContent = engineLabel;
  if (docState) docState.textContent = 'Saved';
  if (statusState) statusState.textContent = 'Ready';

  inst.store.subscribe(() => {
    projectStatus();
    projectFormatToolbar();
    markDirty();
    refreshZoom();
  });
  projectStatus();
  projectFormatToolbar();
  renderSheetTabs();
  refreshZoom();

  // Reflect Format Painter state on the toolbar button (any path can deactivate
  // it — Esc, post-paint, or programmatic).
  inst.formatPainter?.subscribe((active, sticky) => {
    formatPainterBtn?.classList.toggle(ACTIVE_CLASS, active);
    formatPainterBtn?.classList.toggle('app__tool--sticky', active && sticky);
  });
}

document.getElementById('btn-autosum')?.addEventListener('click', () => {
  if (!inst) return;
  const result = autoSum(inst.store.getState(), inst.workbook);
  if (!result) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(result.addr.sheet));
  mutators.setActive(inst.store, result.addr);
  (sheetEl as HTMLElement).focus();
});
document.getElementById('btn-pivot')?.addEventListener('click', () => {
  inst?.openPivotTableDialog();
});
document.getElementById('btn-hyperlink')?.addEventListener('click', () => {
  inst?.openHyperlinkDialog();
});
const openCommentDialog = (): void => {
  inst?.openCommentDialog();
};
document.getElementById('btn-comment')?.addEventListener('click', openCommentDialog);
document.getElementById('btn-review-comment')?.addEventListener('click', openCommentDialog);
document.getElementById('btn-help-readme')?.addEventListener('click', () => {
  window.open('https://github.com/libraz/formulon-cell#readme', '_blank', 'noopener,noreferrer');
});

document.getElementById('btn-undo')?.addEventListener('click', () => {
  if (!inst) return;
  if (!inst.undo()) return;
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-redo')?.addEventListener('click', () => {
  if (!inst) return;
  if (!inst.redo()) return;
  (sheetEl as HTMLElement).focus();
});

// Format Painter — single click arms one-shot, double click arms sticky mode.
// Re-clicking the active button deactivates.
const formatPainterBtn = document.getElementById('btn-format-painter');
let painterStickyTimer: number | null = null;
formatPainterBtn?.addEventListener('click', () => {
  if (!inst) return;
  // Defer one-shot activation briefly so a follow-up click within the
  // dblclick window can promote it to sticky without painting twice.
  if (painterStickyTimer != null) return;
  painterStickyTimer = window.setTimeout(() => {
    painterStickyTimer = null;
    if (!inst) return;
    const fp = inst.formatPainter;
    if (!fp) return;
    if (fp.isActive()) fp.deactivate();
    else fp.activate(false);
    (sheetEl as HTMLElement).focus();
    formatPainterBtn?.classList.toggle(ACTIVE_CLASS, fp.isActive());
  }, 220);
});
formatPainterBtn?.addEventListener('dblclick', () => {
  if (!inst) return;
  if (painterStickyTimer != null) {
    clearTimeout(painterStickyTimer);
    painterStickyTimer = null;
  }
  const fp = inst.formatPainter;
  if (!fp) return;
  fp.activate(true);
  (sheetEl as HTMLElement).focus();
  formatPainterBtn?.classList.toggle(ACTIVE_CLASS, fp.isActive());
});

const wireFormat = (
  id: string,
  fn: (
    state: ReturnType<SpreadsheetInstance['store']['getState']>,
    store: SpreadsheetInstance['store'],
  ) => void,
): void => {
  document.getElementById(id)?.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    // Wrap each toolbar mutation so Cmd+Z reverts the format change.
    recordFormatChange(i.history, i.store, () => {
      fn(i.store.getState(), i.store);
    });
    (sheetEl as HTMLElement).focus();
  });
};

wireFormat('btn-bold', toggleBold);
wireFormat('btn-italic', toggleItalic);
wireFormat('btn-underline', toggleUnderline);
wireFormat('btn-strike', toggleStrike);
wireFormat('btn-currency', cycleCurrency);
wireFormat('btn-percent', cyclePercent);
wireFormat('btn-comma', (state, store) => setNumFmt(state, store, { kind: 'fixed', decimals: 2 }));
wireFormat('btn-font-grow', (state, store) => {
  const a = state.selection.active;
  const f = state.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  setFont(state, store, { fontSize: (f?.fontSize ?? 11) + 1 });
});
wireFormat('btn-font-shrink', (state, store) => {
  const a = state.selection.active;
  const f = state.format.formats.get(`${a.sheet}:${a.row}:${a.col}`);
  setFont(state, store, { fontSize: Math.max(1, (f?.fontSize ?? 11) - 1) });
});
wireFormat('btn-align-left', (state, store) => setAlign(state, store, 'left'));
wireFormat('btn-align-center', (state, store) => setAlign(state, store, 'center'));
wireFormat('btn-align-right', (state, store) => setAlign(state, store, 'right'));
wireFormat('btn-top', (state, store) => setVAlign(state, store, 'top'));
wireFormat('btn-middle', (state, store) => setVAlign(state, store, 'middle'));
wireFormat('btn-decimals-up', (state, store) => bumpDecimals(state, store, 1));
wireFormat('btn-decimals-down', (state, store) => bumpDecimals(state, store, -1));

void clearFormat; // Reserved for a "Clear formatting" menu item; keep the import live.

// ── Borders dropdown (Excel-365 parity) ──────────────────────────────────
// Wires the integrated dropdown built by createBordersMenu(): edge / frame
// / combined presets, "More borders..." entry, and the "罫線の作成" block
// which arms the border-draw controller (drives pointer-edge editing on
// the grid) and exposes two submenus for the line color & line style
// brush settings.
const borderBtn = document.getElementById('btn-borders');
const borderMenu = document.getElementById('menu-borders');
const lineColorSubmenu =
  borderMenu?.querySelector<HTMLElement>('.app__submenu--line-color') ?? null;
const lineStyleSubmenu =
  borderMenu?.querySelector<HTMLElement>('.app__submenu--line-style') ?? null;

const BORDER_DRAW_ACTIVE_CLASS = 'app__menu-item--active';

const closeBorderSubmenus = (): void => {
  if (lineColorSubmenu) lineColorSubmenu.hidden = true;
  if (lineStyleSubmenu) lineStyleSubmenu.hidden = true;
  borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]').forEach((b) => {
    b.setAttribute('aria-expanded', 'false');
  });
};

const closeBorderMenu = (restoreFocus = false): void => {
  if (!borderMenu) return;
  borderMenu.hidden = true;
  borderBtn?.setAttribute('aria-expanded', 'false');
  closeBorderSubmenus();
  if (restoreFocus) borderBtn?.focus();
};

const refreshBorderMenuState = (): void => {
  if (!borderMenu) return;
  // Reflect currently-armed draw mode in the menu so the user can see
  // (and toggle off) the active brush.
  const mode = inst?.borderDraw?.getMode() ?? null;
  borderMenu.querySelectorAll<HTMLButtonElement>('[data-border-draw]').forEach((btn) => {
    const armed = btn.dataset.borderDraw === mode;
    btn.classList.toggle(BORDER_DRAW_ACTIVE_CLASS, armed);
    btn.setAttribute('aria-checked', armed ? 'true' : 'false');
  });
};

const openBorderMenu = (): void => {
  if (!borderMenu) return;
  refreshBorderMenuState();
  borderMenu.hidden = false;
  borderBtn?.setAttribute('aria-expanded', 'true');
  focusMenuItem(borderMenu);
};

borderBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!borderMenu) return;
  if (borderMenu.hidden) openBorderMenu();
  else closeBorderMenu();
});

document.addEventListener('mousedown', (e) => {
  if (!borderMenu || borderMenu.hidden) return;
  if (borderMenu.contains(e.target as Node)) return;
  if (borderBtn?.contains(e.target as Node)) return;
  closeBorderMenu();
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !borderMenu?.hidden) closeBorderMenu(true);
});

borderMenu?.addEventListener('keydown', (e) => {
  handleMenuKeydown(e, borderMenu, { close: closeBorderMenu, restoreFocusTo: borderBtn });
});

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

// Map menu key → engine preset. `clear` is the "罫線なし" entry: the
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

borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-preset]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    const key = btn.dataset.borderPreset ?? '';
    if (key === 'format') {
      closeBorderMenu();
      i.openFormatDialog();
      return;
    }
    const preset = MENU_TO_PRESET[key];
    if (!preset) return;
    closeBorderMenu();
    i.borderDraw?.deactivate();
    applyRibbonFormat((state, store) => setBorderPreset(state, store, preset, selectedBorderStyle));
  });
});

borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-draw]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    const action = btn.dataset.borderDraw as 'draw' | 'grid' | 'erase' | undefined;
    if (!action) return;
    const draw = i.borderDraw;
    if (!draw) return;
    if (draw.getMode() === action) {
      draw.deactivate();
    } else {
      draw.activate(action, selectedBorderStyle, selectedBorderColor);
    }
    closeBorderMenu();
    refreshBorderMenuState();
    (sheetEl as HTMLElement).focus();
  });
});

const openSubmenu = (which: 'lineColor' | 'lineStyle'): void => {
  if (which === 'lineColor') {
    if (lineStyleSubmenu) lineStyleSubmenu.hidden = true;
    if (lineColorSubmenu) lineColorSubmenu.hidden = false;
  } else {
    if (lineColorSubmenu) lineColorSubmenu.hidden = true;
    if (lineStyleSubmenu) lineStyleSubmenu.hidden = false;
  }
  borderMenu?.querySelectorAll<HTMLButtonElement>('[data-border-submenu]').forEach((b) => {
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
// Excel's single-active-submenu behavior.
borderMenu
  ?.querySelectorAll<HTMLButtonElement>('[data-border-preset], [data-border-draw]')
  .forEach((btn) => {
    btn.addEventListener('mouseenter', closeBorderSubmenus);
  });

lineColorSubmenu
  ?.querySelectorAll<HTMLButtonElement>('[data-border-line-color]')
  .forEach((swatch) => {
    swatch.addEventListener('click', () => {
      const hex = swatch.dataset.borderLineColor ?? '#000000';
      selectedBorderColor = hex;
      lineColorSubmenu
        .querySelectorAll<HTMLButtonElement>('[data-border-line-color]')
        .forEach((s) => {
          s.setAttribute('aria-checked', s === swatch ? 'true' : 'false');
        });
      inst?.borderDraw?.setColor(hex);
      closeBorderSubmenus();
    });
  });

lineStyleSubmenu
  ?.querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
  .forEach((styleBtn) => {
    styleBtn.addEventListener('click', () => {
      const value = styleBtn.dataset.borderLineStyle ?? 'thin';
      if (value !== 'none') {
        selectedBorderStyle = value as CellBorderStyle;
        inst?.borderDraw?.setStyle(selectedBorderStyle);
      }
      lineStyleSubmenu
        .querySelectorAll<HTMLButtonElement>('[data-border-line-style]')
        .forEach((s) => {
          s.setAttribute('aria-checked', s === styleBtn ? 'true' : 'false');
        });
      closeBorderSubmenus();
    });
  });

// ── Freeze Panes menu ─────────────────────────────────────────────────────
const freezeBtn = document.getElementById('btn-freeze');
const freezeMenu = document.getElementById('menu-freeze');

const closeFreezeMenu = (restoreFocus = false): void => {
  if (!freezeMenu) return;
  freezeMenu.hidden = true;
  freezeBtn?.setAttribute('aria-expanded', 'false');
  if (restoreFocus) freezeBtn?.focus();
};
const openFreezeMenu = (): void => {
  if (!freezeMenu) return;
  freezeMenu.hidden = false;
  freezeBtn?.setAttribute('aria-expanded', 'true');
  focusMenuItem(freezeMenu);
};

freezeBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!freezeMenu) return;
  if (freezeMenu.hidden) openFreezeMenu();
  else closeFreezeMenu();
});

document.addEventListener('mousedown', (e) => {
  if (!freezeMenu || freezeMenu.hidden) return;
  if (freezeMenu.contains(e.target as Node)) return;
  if (freezeBtn?.contains(e.target as Node)) return;
  closeFreezeMenu();
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !freezeMenu?.hidden) closeFreezeMenu(true);
});

freezeMenu?.addEventListener('keydown', (e) => {
  handleMenuKeydown(e, freezeMenu, { close: closeFreezeMenu, restoreFocusTo: freezeBtn });
});

freezeMenu?.querySelectorAll<HTMLButtonElement>('[data-freeze]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const i = inst;
    if (!i) return;
    const action = btn.dataset.freeze;
    const s = i.store.getState();

    let rows = s.layout.freezeRows;
    let cols = s.layout.freezeCols;
    if (action === 'row') {
      rows = 1;
      cols = 0;
    } else if (action === 'col') {
      rows = 0;
      cols = 1;
    } else if (action === 'selection') {
      // Freeze rows above and columns left of the active cell.
      rows = s.selection.active.row;
      cols = s.selection.active.col;
    } else if (action === 'off') {
      rows = 0;
      cols = 0;
    }

    setFreezePanes(i.store, i.history, rows, cols, i.workbook);
    closeFreezeMenu();
    (sheetEl as HTMLElement).focus();
  });
});

themeToggle?.addEventListener('click', () => {
  uiTheme = uiTheme === 'light' ? 'dark' : 'light';
  html.dataset.theme = uiTheme;
  if (themeLabel) themeLabel.textContent = uiTheme === 'light' ? 'Light' : 'Dark';
  themeToggle.setAttribute('aria-pressed', uiTheme === 'dark' ? 'true' : 'false');
  // Theme is a UI-only preference; don't let the resulting store update mark the workbook as edited.
  suppressDirty = true;
  inst?.setTheme(toCore(uiTheme));
  suppressDirty = false;
});

// ── File menu (New / Open / Save / Save As) ───────────────────────────────
const fileMenuBtn = document.getElementById('menu-file');
const fileMenuDrop = document.getElementById('menu-file-dropdown');
const fileInput = document.getElementById('file-input') as HTMLInputElement | null;

let docName = 'Untitled';

const setDocName = (name: string): void => {
  docName = name;
  const el = document.getElementById('doc-name');
  if (el) el.textContent = name;
};

const openFileMenu = (): void => {
  if (!fileMenuDrop) return;
  fileMenuDrop.hidden = false;
  fileMenuBtn?.setAttribute('aria-expanded', 'true');
};
const closeFileMenu = (): void => {
  if (!fileMenuDrop) return;
  fileMenuDrop.hidden = true;
  fileMenuBtn?.setAttribute('aria-expanded', 'false');
};

fileMenuBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!fileMenuDrop) return;
  if (fileMenuDrop.hidden) openFileMenu();
  else closeFileMenu();
});

document.addEventListener('mousedown', (e) => {
  if (!fileMenuDrop || fileMenuDrop.hidden) return;
  if (fileMenuDrop.contains(e.target as Node)) return;
  if (fileMenuBtn?.contains(e.target as Node)) return;
  closeFileMenu();
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !fileMenuDrop?.hidden) closeFileMenu();
});

const triggerOpen = (): void => fileInput?.click();

const downloadBytes = (bytes: Uint8Array, filename: string): void => {
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
};

const triggerSave = (filename = `${docName.replace(/\.xlsx$/i, '')}.xlsx`): void => {
  if (!inst) return;
  try {
    const bytes = inst.workbook.save();
    downloadBytes(bytes, filename);
    if (docState) docState.textContent = 'Saved';
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('save failed', err);
    if (docState) docState.textContent = 'Save failed';
  }
};

const triggerSaveAs = async (): Promise<void> => {
  const name = await showPrompt({
    title: 'Save As',
    label: 'File name',
    initial: docName,
    okLabel: 'Save',
    validate: (value) => (value.trim() ? null : 'Enter a file name.'),
  });
  if (!name) return;
  const trimmed = name.trim();
  setDocName(trimmed);
  triggerSave(trimmed.endsWith('.xlsx') ? trimmed : `${trimmed}.xlsx`);
};

const loadXlsxFile = async (file: File): Promise<void> => {
  if (!inst) return;
  if (docState) docState.textContent = 'Loading…';
  try {
    const buf = await file.arrayBuffer();
    const next = await WorkbookHandle.loadBytes(new Uint8Array(buf));
    await inst.setWorkbook(next);
    setDocName(file.name);
    if (docState) docState.textContent = 'Saved';
    renderSheetTabs();
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('open failed', err);
    if (docState) docState.textContent = 'Open failed';
    void showMessage({
      title: 'Open failed',
      message: err instanceof Error ? err.message : String(err),
    });
  }
};

fileInput?.addEventListener('change', () => {
  const f = fileInput.files?.[0];
  if (f) void loadXlsxFile(f);
  fileInput.value = ''; // allow same-file re-open
});

fileMenuDrop?.querySelectorAll<HTMLButtonElement>('[data-file]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const action = btn.dataset.file;
    closeFileMenu();
    if (!inst) return;
    if (action === 'new') {
      void (async () => {
        const next = await WorkbookHandle.createDefault();
        await inst?.setWorkbook(next);
        setDocName('Untitled');
        if (docState) docState.textContent = 'Saved';
        renderSheetTabs();
      })();
    } else if (action === 'open') {
      triggerOpen();
    } else if (action === 'save') {
      triggerSave();
    } else if (action === 'save-as') {
      void triggerSaveAs();
    }
  });
});

// Drag & drop xlsx onto the page.
window.addEventListener('dragover', (e) => {
  if (!e.dataTransfer) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  const f = e.dataTransfer?.files?.[0];
  if (!f) return;
  if (!/\.xlsx?$/i.test(f.name)) return;
  void loadXlsxFile(f);
});

// Ctrl/Cmd-O / Ctrl/Cmd-S / Ctrl/Cmd-N for file actions.
window.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (k === 'o') {
    e.preventDefault();
    triggerOpen();
  } else if (k === 's') {
    e.preventDefault();
    if (e.shiftKey) void triggerSaveAs();
    else triggerSave();
  } else if (k === 'n' && !e.shiftKey) {
    // Ctrl+N — create a fresh workbook in place.
    e.preventDefault();
    void (async () => {
      const next = await WorkbookHandle.createDefault();
      await inst?.setWorkbook(next);
      setDocName('Untitled');
      renderSheetTabs();
    })();
  }
});

// Mark the document dirty whenever any cell change flows through.
let dirtyTimer: number | null = null;
let suppressDirty = false;
const markDirty = (): void => {
  if (suppressDirty) return;
  if (dirtyTimer != null) return;
  dirtyTimer = window.setTimeout(() => {
    dirtyTimer = null;
    if (docState) docState.textContent = 'Edited';
  }, 200);
};
// Subscribe once boot completes — see end of boot().

const refreshWorkbookCells = (): void => {
  if (!inst) return;
  mutators.replaceCells(inst.store, inst.workbook.cells(inst.store.getState().data.sheetIndex));
};

const focusSheet = (): void => {
  (sheetEl as HTMLElement).focus();
};

const selectedRowCount = (): number => {
  if (!inst) return 1;
  const r = inst.store.getState().selection.range;
  return Math.max(1, r.r1 - r.r0 + 1);
};

const selectedColCount = (): number => {
  if (!inst) return 1;
  const r = inst.store.getState().selection.range;
  return Math.max(1, r.c1 - r.c0 + 1);
};

const openFilterForSelection = (): void => {
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  mutators.setFilterRange(inst.store, r);
  const sheetRect = (sheetEl as HTMLElement).getBoundingClientRect();
  filterDropdown?.open(r, r.c0, { x: sheetRect.left + 80, y: sheetRect.top, h: 32 });
  focusSheet();
};

const sortSelection = (direction: 'asc' | 'desc'): void => {
  if (!inst) return;
  const state = inst.store.getState();
  const r = state.selection.range;
  if (r.r0 === r.r1 && r.c0 === r.c1) return;
  sortRange(state, inst.store, inst.workbook, r, { byCol: r.c0, direction });
  refreshWorkbookCells();
  focusSheet();
};

const removeDuplicateRows = (): void => {
  if (!inst) return;
  const state = inst.store.getState();
  const removed = removeDuplicates(state, inst.store, inst.workbook, state.selection.range);
  refreshWorkbookCells();
  if (statusMetric)
    statusMetric.textContent = `Removed ${removed} duplicate row${removed === 1 ? '' : 's'}`;
  focusSheet();
};

const createTableFromSelection = (): void => {
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  formatAsTable(inst.store, r);
  focusSheet();
};

const createChartFromSelection = (): void => {
  if (!inst) return;
  const r = inst.store.getState().selection.range;
  const count = inst.store.getState().charts.charts.length;
  createSessionChart(inst.store, r, {
    id: `ribbon-chart-${r.sheet}-${r.r0}-${r.c0}-${r.r1}-${r.c1}-${count}`,
    kind: 'column',
    title: null,
    x: 340 + (count % 3) * 24,
    y: 96 + (count % 3) * 24,
    w: 360,
    h: 220,
  });
  focusSheet();
};

const copySelectionToClipboard = async (): Promise<void> => {
  if (!inst) return;
  const result = copy(inst.store.getState());
  if (!result) return;
  await navigator.clipboard?.writeText(result.tsv);
  focusSheet();
};

const cutSelectionToClipboard = async (): Promise<void> => {
  if (!inst) return;
  const result = cut(inst.store.getState(), inst.workbook);
  if (!result) return;
  await navigator.clipboard?.writeText(result.tsv);
  refreshWorkbookCells();
  focusSheet();
};

const pasteClipboardIntoSelection = async (): Promise<void> => {
  if (!inst) return;
  const text = await navigator.clipboard?.readText();
  if (!text) return;
  pasteTSV(inst.store.getState(), inst.workbook, text);
  refreshWorkbookCells();
  focusSheet();
};

const addrLabel = (row: number, col: number): string => `${colLabel(col)}${row + 1}`;

const reviewCellsForSheet = (sheet: number): ReviewCell[] => {
  if (!inst) return [];
  return Array.from(inst.workbook.cells(sheet), (entry) => ({
    label: addrLabel(entry.addr.row, entry.addr.col),
    value:
      entry.value.kind === 'text'
        ? { kind: 'text' as const, value: entry.value.value }
        : entry.value.kind === 'error'
          ? { kind: 'error' as const, text: entry.value.text }
          : entry.value.kind === 'number'
            ? { kind: 'number' as const }
            : entry.value.kind === 'bool'
              ? { kind: 'bool' as const }
              : { kind: 'blank' as const },
    formula: entry.formula,
  }));
};

const showRibbonReport = (title: string, items: readonly RibbonReportItem[]): void => {
  void showReport({ title, items });
};

const selectedPlainText = (): string => {
  if (!inst) return '';
  const state = inst.store.getState();
  const range = state.selection.range;
  const lines: string[] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    const cells: string[] = [];
    for (let col = range.c0; col <= range.c1; col += 1) {
      const formula = inst.workbook.cellFormula({ sheet: range.sheet, row, col });
      if (formula) {
        cells.push(formula);
        continue;
      }
      const value = inst.workbook.getValue({ sheet: range.sheet, row, col });
      if (value.kind === 'text') cells.push(value.value);
      else if (value.kind === 'number') cells.push(String(value.value));
      else if (value.kind === 'bool') cells.push(value.value ? 'TRUE' : 'FALSE');
      else if (value.kind === 'error') cells.push(value.text);
      else cells.push('');
    }
    lines.push(cells.join('\t'));
  }
  return lines.join('\n').trim();
};

const runAccessibilityCheck = (): void => {
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const items = analyzeAccessibilityCells(reviewCellsForSheet(sheet));
  if (statusMetric)
    statusMetric.textContent = `Accessibility · ${items.filter((i) => i.severity === 'warning').length} warnings`;
  showRibbonReport('Accessibility Check', items);
};

const runSpellingReview = (): void => {
  if (!inst) return;
  const sheet = inst.store.getState().data.sheetIndex;
  const items = analyzeSpellingCells(reviewCellsForSheet(sheet));
  if (statusMetric)
    statusMetric.textContent = `Spelling · ${items.filter((i) => i.severity === 'warning').length} warnings`;
  showRibbonReport('Spelling Review', items);
};

const openTranslateReview = (): void => {
  const text = selectedPlainText();
  showRibbonReport('Translate Selection', [
    text
      ? {
          severity: 'info',
          label: 'Selection text',
          detail: text.length > 500 ? `${text.slice(0, 500)}...` : text,
        }
      : {
          severity: 'info',
          label: 'No text selected',
          detail: 'Select cells containing text before using Translate.',
        },
    {
      severity: 'info',
      label: 'Privacy',
      detail: 'No text is sent to an external translation service from the playground.',
    },
  ]);
};

const runPlaygroundScript = async (): Promise<void> => {
  if (!inst) return;
  const command = await showPrompt({
    title: 'Run Script',
    label: 'Command',
    placeholder: 'uppercase, lowercase, trim, clear',
    okLabel: 'Run',
    validate: (value) =>
      parseScriptCommand(value) ? null : 'Use one of: uppercase, lowercase, trim, clear.',
  });
  if (!command || !inst) return;
  const op = parseScriptCommand(command);
  if (!op) return;
  const range = inst.store.getState().selection.range;
  let changed = 0;
  inst.history.begin();
  try {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        const value = inst.workbook.getValue(addr);
        if (op === 'clear') {
          if (value.kind !== 'blank' || inst.workbook.cellFormula(addr)) {
            inst.workbook.setBlank(addr);
            changed += 1;
          }
          continue;
        }
        if (value.kind !== 'text') continue;
        const next = applyTextScript(value.value, op);
        if (next !== value.value) {
          inst.workbook.setText(addr, next);
          changed += 1;
        }
      }
    }
  } finally {
    inst.history.end();
  }
  refreshWorkbookCells();
  if (statusMetric)
    statusMetric.textContent = `Script · ${changed} cell${changed === 1 ? '' : 's'}`;
  focusSheet();
};

const openAddInManager = (): void => {
  showRibbonReport('Add-ins', [
    {
      severity: 'info',
      label: 'Built-in add-ins',
      detail:
        'Charts, PivotTable dialog, Watch Window, and PDF/Print are available in this playground.',
    },
    {
      severity: 'info',
      label: 'External add-ins',
      detail: 'External add-in packages are not loaded automatically in the playground.',
    },
  ]);
};

const applyRibbonCommand = (id: string): boolean => {
  const i = inst;
  if (!i) return false;
  const state = i.store.getState();
  const range = state.selection.range;
  switch (id) {
    case 'pageSetup':
    case 'pageSetupAdvanced':
      i.openPageSetup();
      return true;
    case 'print':
    case 'printPageLayout':
    case 'pdf':
      i.print();
      return true;
    case 'links':
    case 'linksInsert':
    case 'linksData':
      i.openExternalLinksDialog();
      return true;
    case 'formatCells':
    case 'formatCellsHome':
      i.openFormatDialog();
      return true;
    case 'gotoSpecial':
    case 'gotoSpecialHome':
      i.openGoToSpecial();
      return true;
    case 'paste':
      void pasteClipboardIntoSelection();
      return true;
    case 'cut':
      void cutSelectionToClipboard();
      return true;
    case 'copy':
      void copySelectionToClipboard();
      return true;
    case 'clearFormat':
      applyRibbonFormat((s, store) => clearFormat(s, store));
      return true;
    case 'general':
      applyRibbonFormat((s, store) => setNumFmt(s, store, { kind: 'general' }));
      return true;
    case 'conditional':
      i.openConditionalDialog();
      return true;
    case 'cellStyles':
      i.openCellStylesGallery();
      return true;
    case 'rules':
      i.openCfRulesDialog();
      return true;
    case 'insertRows':
      insertRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'deleteRows':
      deleteRows(i.store, i.workbook, i.history, range.r0, selectedRowCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'insertCols':
      insertCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'deleteCols':
      deleteCols(i.store, i.workbook, i.history, range.c0, selectedColCount());
      refreshWorkbookCells();
      focusSheet();
      return true;
    case 'sortAscHome':
    case 'sortAsc':
      sortSelection('asc');
      return true;
    case 'sortDesc':
      sortSelection('desc');
      return true;
    case 'filterHome':
      openFilterForSelection();
      return true;
    case 'drawPen':
      applyRibbonFormat((s, store) => setBorderPreset(s, store, 'all', selectedBorderStyle));
      return true;
    case 'drawErase':
      applyRibbonFormat((s, store) => setBorderPreset(s, store, 'none', selectedBorderStyle));
      return true;
    case 'findHome':
    case 'findReview':
      i.openFindReplace();
      return true;
    case 'spellingReview':
      runSpellingReview();
      return true;
    case 'translateReview':
      openTranslateReview();
      return true;
    case 'accessibility':
      runAccessibilityCheck();
      return true;
    case 'formatTableInsert':
      createTableFromSelection();
      return true;
    case 'namedRangesInsert':
    case 'namedRanges':
      i.openNamedRangeDialog();
      return true;
    case 'removeDupesInsert':
    case 'removeDupes':
      removeDuplicateRows();
      return true;
    case 'chartInsert':
      createChartFromSelection();
      return true;
    case 'fxInsert':
    case 'fx':
      i.openFunctionArguments();
      return true;
    case 'autosumFormula': {
      const result = autoSum(i.store.getState(), i.workbook);
      if (result) {
        refreshWorkbookCells();
        mutators.setActive(i.store, result.addr);
      }
      focusSheet();
      return true;
    }
    case 'sum':
      i.openFunctionArguments('SUM');
      return true;
    case 'avg':
      i.openFunctionArguments('AVERAGE');
      return true;
    case 'precedents':
      i.tracePrecedents();
      return true;
    case 'dependents':
      i.traceDependents();
      return true;
    case 'clearArrows':
      i.clearTraces();
      return true;
    case 'recalcNow':
      i.recalc();
      focusSheet();
      return true;
    case 'calcOptions':
      i.openIterativeDialog();
      return true;
    case 'watch':
    case 'watchView':
      i.toggleWatchWindow();
      return true;
    case 'hideRows':
      hideRows(i.store, i.history, range.r0, range.r1, i.workbook);
      focusSheet();
      return true;
    case 'hideCols':
      hideCols(i.store, i.history, range.c0, range.c1, i.workbook);
      focusSheet();
      return true;
    case 'newCommentReview':
      i.openCommentDialog();
      return true;
    case 'protectReview':
    case 'protect':
      i.toggleSheetProtection();
      focusSheet();
      return true;
    case 'script':
      void runPlaygroundScript();
      return true;
    case 'addIn':
      openAddInManager();
      return true;
    case 'zoom75':
      setSheetZoom(i.store, 0.75, i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    case 'zoom100':
      setSheetZoom(i.store, 1, i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    case 'zoom125':
      setSheetZoom(i.store, 1.25, i.workbook);
      refreshZoom();
      focusSheet();
      return true;
    default:
      return false;
  }
};

// ── Ribbon tab strip ────────────────────────────────────────────────────
const selectRibbonTab = (tabId: RibbonTab, focusTab = false): void => {
  if (!ribbonRoot) return;
  activeRibbonTab = tabId;
  for (const item of ribbonRoot.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
    const isActive = item.dataset.ribbonTab === activeRibbonTab;
    item.classList.toggle('demo__ribbon-tab--active', isActive);
    item.setAttribute('aria-selected', isActive ? 'true' : 'false');
    item.tabIndex = isActive ? 0 : -1;
    if (focusTab && isActive) item.focus({ preventScroll: true });
  }
  for (const panel of ribbonRoot.querySelectorAll<HTMLElement>('[data-ribbon-panel]')) {
    panel.hidden = panel.dataset.ribbonPanel !== activeRibbonTab;
  }
};

ribbonRoot?.addEventListener('click', (event) => {
  const tab = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!tab) return;
  selectRibbonTab((tab.dataset.ribbonTab as RibbonTab | undefined) ?? 'home');
});

ribbonRoot?.addEventListener('keydown', (event) => {
  const tab = (event.target as Element | null)?.closest<HTMLButtonElement>('[data-ribbon-tab]');
  if (!tab) return;
  const tabs = Array.from(ribbonRoot.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]'));
  const current = Math.max(0, tabs.indexOf(tab));
  let next = current;
  if (event.key === 'ArrowRight') next = (current + 1) % tabs.length;
  else if (event.key === 'ArrowLeft') next = (current - 1 + tabs.length) % tabs.length;
  else if (event.key === 'Home') next = 0;
  else if (event.key === 'End') next = tabs.length - 1;
  else return;
  event.preventDefault();
  const nextTab = tabs[next]?.dataset.ribbonTab as RibbonTab | undefined;
  if (nextTab) selectRibbonTab(nextTab, true);
});

ribbonRoot?.addEventListener('click', (event) => {
  const button = (event.target as Element | null)?.closest<HTMLButtonElement>(
    'button[data-ribbon-command]',
  );
  if (!button || button.disabled) return;
  const id = button.dataset.ribbonCommand;
  if (!id) return;
  if (legacyCommandIds[id]) return;
  if (applyRibbonCommand(id)) {
    event.preventDefault();
    event.stopPropagation();
  }
});

// ── View menu (Show Formulas / R1C1 / Grid / Headers toggles) ────────────
const viewBtn = document.getElementById('menu-view');
const viewDrop = document.getElementById('menu-view-dropdown');
const closeViewMenu = (): void => {
  if (!viewDrop) return;
  viewDrop.hidden = true;
  viewBtn?.setAttribute('aria-expanded', 'false');
};
const refreshViewMenu = (): void => {
  if (!inst || !viewDrop) return;
  const ui = inst.store.getState().ui;
  const update = (action: string, on: boolean): void => {
    const item = viewDrop.querySelector<HTMLElement>(`[data-view="${action}"] [data-fc-check]`);
    if (item) item.textContent = on ? '✓' : '';
  };
  update('show-formulas', !!ui.showFormulas);
  update('r1c1', !!ui.r1c1);
  update('grid', ui.showGridLines !== false);
  update('headers', ui.showHeaders !== false);
};
viewBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!viewDrop) return;
  refreshViewMenu();
  viewDrop.hidden = !viewDrop.hidden;
  viewBtn.setAttribute('aria-expanded', viewDrop.hidden ? 'false' : 'true');
});
document.addEventListener('mousedown', (e) => {
  if (!viewDrop || viewDrop.hidden) return;
  if (viewDrop.contains(e.target as Node) || viewBtn?.contains(e.target as Node)) return;
  closeViewMenu();
});
viewDrop?.querySelectorAll<HTMLButtonElement>('[data-view]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.view;
    const ui = inst.store.getState().ui;
    if (action === 'show-formulas') mutators.setShowFormulas(inst.store, !ui.showFormulas);
    else if (action === 'r1c1') mutators.setR1C1(inst.store, !ui.r1c1);
    else if (action === 'grid') mutators.setShowGridLines(inst.store, !ui.showGridLines);
    else if (action === 'headers') mutators.setShowHeaders(inst.store, !ui.showHeaders);
    refreshViewMenu();
  });
});

// ── Tools menu (Iterative / Names / Conditional) ─────────────────────────
const toolsBtn = document.getElementById('menu-tools');
const toolsDrop = document.getElementById('menu-tools-dropdown');
const closeToolsMenu = (): void => {
  if (!toolsDrop) return;
  toolsDrop.hidden = true;
  toolsBtn?.setAttribute('aria-expanded', 'false');
};
toolsBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  if (!toolsDrop) return;
  toolsDrop.hidden = !toolsDrop.hidden;
  toolsBtn.setAttribute('aria-expanded', toolsDrop.hidden ? 'false' : 'true');
});
document.addEventListener('mousedown', (e) => {
  if (!toolsDrop || toolsDrop.hidden) return;
  if (toolsDrop.contains(e.target as Node) || toolsBtn?.contains(e.target as Node)) return;
  closeToolsMenu();
});
toolsDrop?.querySelectorAll<HTMLButtonElement>('[data-tools]').forEach((btn) => {
  btn.addEventListener('click', () => {
    if (!inst) return;
    const action = btn.dataset.tools;
    closeToolsMenu();
    if (action === 'iterative') inst.openIterativeDialog();
    else if (action === 'named') inst.openNamedRangeDialog();
    else if (action === 'conditional') inst.openConditionalDialog();
  });
});

// ── Sheet tabs ───────────────────────────────────────────────────────────
const tabsList = document.getElementById('sheet-tabs');
const tabAddBtn = document.getElementById('btn-sheet-add');
const tabPrevBtn = document.getElementById('btn-sheet-prev');
const tabNextBtn = document.getElementById('btn-sheet-next');

const renderSheetTabs = (): void => {
  if (!inst || !tabsList) return;
  const wb = inst.workbook;
  const state = inst.store.getState();
  const activeIdx = state.data.sheetIndex;
  const hidden = state.layout.hiddenSheets;
  const n = wb.sheetCount;
  tabsList.replaceChildren();
  for (let i = 0; i < n; i += 1) {
    if (hidden.has(i)) continue;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'app__tab';
    if (i === activeIdx) btn.classList.add('app__tab--active');
    btn.setAttribute('role', 'tab');
    btn.setAttribute('aria-selected', i === activeIdx ? 'true' : 'false');
    const label = document.createElement('span');
    label.className = 'app__tab-label';
    label.textContent = wb.sheetName(i);
    btn.appendChild(label);
    btn.addEventListener('click', () => switchSheet(i));
    btn.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      openTabMenu(i, e.clientX, e.clientY);
    });
    tabsList.appendChild(btn);
  }
  // "Unhide…" affordance — surfaced as an extra tab pill when at least one
  // sheet is hidden. Click opens a list of hidden sheets to restore.
  if (hidden.size > 0) {
    const unhide = document.createElement('button');
    unhide.type = 'button';
    unhide.className = 'app__tab app__tab--unhide';
    unhide.textContent = `Unhide… (${hidden.size})`;
    unhide.addEventListener('click', (e) => {
      const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
      openUnhideMenu(r.left, r.bottom);
    });
    tabsList.appendChild(unhide);
  }
};

const openUnhideMenu = (x: number, y: number): void => {
  if (!inst) return;
  closeTabMenu();
  const wb = inst.workbook;
  const store = inst.store;
  const hidden = store.getState().layout.hiddenSheets;
  if (hidden.size === 0) return;

  const menu = document.createElement('div');
  menu.className = 'app__menu';
  prepareMenu(menu, 'Unhide sheet');
  menu.style.position = 'fixed';
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  menu.style.zIndex = '90';
  let cleanupMenuListeners = (): void => {};

  for (const i of Array.from(hidden).sort((a, b) => a - b)) {
    const it = document.createElement('button');
    it.type = 'button';
    it.className = 'app__menu-item';
    it.setAttribute('role', 'menuitem');
    it.tabIndex = -1;
    it.textContent = wb.sheetName(i);
    it.addEventListener('click', () => {
      closeTabMenu();
      cleanupMenuListeners();
      if (setSheetHidden(store, wb, inst?.history ?? null, i, false)) {
        renderSheetTabs();
      }
    });
    menu.appendChild(it);
  }

  document.body.appendChild(menu);
  tabMenuEl = menu;
  focusMenuItem(menu);

  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) {
    menu.style.left = `${Math.max(0, window.innerWidth - rect.width - 4)}px`;
  }
  if (rect.bottom > window.innerHeight) {
    menu.style.top = `${Math.max(0, window.innerHeight - rect.height - 4)}px`;
  }

  const onDocDown = (ev: MouseEvent): void => {
    if (!tabMenuEl) return;
    if (ev.target instanceof Node && tabMenuEl.contains(ev.target)) return;
    closeTabMenu();
    cleanupMenuListeners();
  };
  const onDocKey = (ev: KeyboardEvent): void => {
    handleMenuKeydown(ev, menu, {
      close: (restoreFocus) => {
        closeTabMenu();
        cleanupMenuListeners();
        if (restoreFocus) {
          document.querySelector<HTMLButtonElement>('.app__tab--unhide')?.focus();
        }
      },
    });
  };
  cleanupMenuListeners = () => {
    document.removeEventListener('mousedown', onDocDown, true);
    document.removeEventListener('keydown', onDocKey, true);
  };
  document.addEventListener('mousedown', onDocDown, true);
  document.addEventListener('keydown', onDocKey, true);
};

let tabMenuEl: HTMLDivElement | null = null;
const closeTabMenu = (): void => {
  if (!tabMenuEl) return;
  tabMenuEl.remove();
  tabMenuEl = null;
};

const openTabMenu = (idx: number, x: number, y: number): void => {
  if (!inst) return;
  openSheetTabMenu({
    closeTabMenu,
    idx,
    inst,
    renderSheetTabs,
    setTabMenuEl: (el) => {
      tabMenuEl = el;
    },
    x,
    y,
  });
};

const switchSheet = (idx: number): void => {
  if (!inst) return;
  const n = inst.workbook.sheetCount;
  if (idx < 0 || idx >= n) return;
  if (inst.store.getState().data.sheetIndex === idx) return;
  mutators.setSheetIndex(inst.store, idx);
  mutators.replaceCells(inst.store, inst.workbook.cells(idx));
  renderSheetTabs();
  (sheetEl as HTMLElement).focus();
};

tabAddBtn?.addEventListener('click', () => {
  if (!inst) return;
  const idx = inst.workbook.addSheet();
  if (idx < 0) return;
  // The wb.subscribe handler in mount.ts will pick up sheet-add as a no-op for cells,
  // but we re-render tabs and switch to the new sheet here.
  renderSheetTabs();
  switchSheet(idx);
});

const { refreshZoom } = setupZoomControls(() => inst);

tabPrevBtn?.addEventListener('click', () => {
  if (!inst) return;
  switchSheet(inst.store.getState().data.sheetIndex - 1);
});
tabNextBtn?.addEventListener('click', () => {
  if (!inst) return;
  switchSheet(inst.store.getState().data.sheetIndex + 1);
});

// ── Merge / Wrap / Sort buttons ───────────────────────────────────────────
document.getElementById('btn-merge')?.addEventListener('click', () => {
  if (!inst) return;
  const s = inst.store.getState();
  const r = s.selection.range;
  const anchorAt0 = s.merges.byAnchor.get(`${r.sheet}:${r.r0}:${r.c0}`);
  const isExactMerge =
    anchorAt0 &&
    r.r0 === anchorAt0.r0 &&
    r.c0 === anchorAt0.c0 &&
    r.r1 === anchorAt0.r1 &&
    r.c1 === anchorAt0.c1;
  if (isExactMerge) applyUnmerge(inst.store, inst.workbook, inst.history, r);
  else applyMerge(inst.store, inst.workbook, inst.history, r);
  (sheetEl as HTMLElement).focus();
});

document.getElementById('btn-wrap')?.addEventListener('click', () => {
  if (!inst) return;
  const current = inst;
  recordFormatChange(inst.history, inst.store, () => {
    toggleWrap(current.store.getState(), current.store);
  });
  (sheetEl as HTMLElement).focus();
});

setupSortMenu({
  getFilterDropdown: () => filterDropdown,
  getInst: () => inst,
  sheetEl: sheetEl as HTMLElement,
  statusMetric,
});

let filterDropdown: ReturnType<typeof attachFilterDropdown> | null = null;

boot().catch((err) => {
  // eslint-disable-next-line no-console
  console.error('formulon-cell boot failed', err);
  if (sheetEl) {
    sheetEl.innerHTML = `<pre style="padding:24px;color:#d24545;font-family:'IBM Plex Mono',monospace;white-space:pre-wrap">${
      err instanceof Error ? (err.stack ?? err.message) : String(err)
    }</pre>`;
  }
});
