import type { CellStyleIncludeOptions } from '../../commands/cell-styles.js';
import type { CellFormat } from '../../store/types.js';
import { appendCheckboxRow } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogNameField,
  appendErrorRow,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface CellStyleDialogResult {
  name: string;
  include: Required<CellStyleIncludeOptions>;
}

interface CellStyleDialogStrings {
  ribbonMenu: {
    cellStyleName: string;
    cellStyleNormal: string;
    cellStyleIncludes: string;
    cellStyleIncludeNumber: string;
    cellStyleIncludeAlignment: string;
    cellStyleIncludeFont: string;
    cellStyleIncludeBorder: string;
    cellStyleIncludeFill: string;
    cellStyleIncludeProtection: string;
  };
  hyperlinkDialog: { cancel: string; ok: string };
  namedRangeDialog: { errorEmptyName: string };
}

type IncludeKey = keyof Required<CellStyleIncludeOptions>;

/** Wording of the desktop Style dialog. */
export interface CellStyleDialogMacStrings {
  title: string;
  nameLabel: string;
  previewLabel: string;
  caption: string;
  propertyColumn: string;
  formatColumn: string;
  properties: Record<IncludeKey, string>;
  formatButton: string;
}

/** Desktop Style dialog: name, preview, a checklist of properties with the format each carries. */
export interface CellStyleDialogMacOptions {
  strings: CellStyleDialogMacStrings;
  /** What the style sets for each property, as text. */
  summaries: Record<IncludeKey, string>;
  /** How the preview paints the style name. */
  preview?: { fill?: string; color?: string; bold?: boolean; italic?: boolean };
  /** Opens Format Cells for the style's formatting. */
  onFormat?: () => void;
}

/** Wording for the one-line format summaries of the desktop Style dialog. */
export interface CellStyleSummaryStrings {
  general: string;
  /** Template with `{vertical}` and `{horizontal}`. */
  alignment: string;
  vertical: { top: string; middle: string; bottom: string };
  horizontal: { general: string; left: string; center: string; right: string };
  noBorder: string;
  border: string;
  noFill: string;
  locked: string;
  unlocked: string;
}

/** What `format` sets for each Style dialog property, as the dialog's summary text. */
export const describeCellStyleFormat = (
  format: CellFormat | undefined,
  t: CellStyleSummaryStrings,
): Record<IncludeKey, string> => {
  const f = format ?? {};
  const horizontal =
    f.align === 'left'
      ? t.horizontal.left
      : f.align === 'center'
        ? t.horizontal.center
        : f.align === 'right'
          ? t.horizontal.right
          : t.horizontal.general;
  const vertical =
    f.vAlign === 'top'
      ? t.vertical.top
      : f.vAlign === 'middle'
        ? t.vertical.middle
        : t.vertical.bottom;
  const font = [f.fontFamily, f.fontSize].filter((part) => part !== undefined).join(' ');
  return {
    number: f.numFmt && f.numFmt.kind !== 'general' ? String(f.numFmt.kind) : t.general,
    alignment: t.alignment.replace('{vertical}', vertical).replace('{horizontal}', horizontal),
    font: f.bold ? `${font} B`.trim() : font,
    border: f.borders && Object.keys(f.borders).length > 0 ? t.border : t.noBorder,
    fill: f.fill ?? t.noFill,
    protection: f.locked === false ? t.unlocked : t.locked,
  };
};

export interface CellStyleDialogOptions {
  title: string;
  strings: CellStyleDialogStrings;
  initialName?: string;
  initialInclude?: CellStyleIncludeOptions;
  /** When given, the dialog takes the desktop Style dialog layout. */
  mac?: CellStyleDialogMacOptions;
}

const DEFAULT_INCLUDE: Required<CellStyleIncludeOptions> = {
  number: true,
  alignment: true,
  font: true,
  border: true,
  fill: true,
  protection: true,
};

export const showCellStyleDialog = (
  opts: CellStyleDialogOptions,
): Promise<CellStyleDialogResult | null> =>
  new Promise<CellStyleDialogResult | null>((resolve) => {
    const t = opts.strings.ribbonMenu;
    const include = { ...DEFAULT_INCLUDE, ...opts.initialInclude };
    const mac = opts.mac;
    const shell = createDialogShell({ title: mac?.strings.title ?? opts.title });
    let nameHost: HTMLElement = shell.body;
    if (mac) {
      shell.panel.classList.add('fc-cellstyledlg');
      nameHost = document.createElement('div');
      nameHost.className = 'fc-cellstyledlg__namerow';
      shell.body.appendChild(nameHost);
    }
    const name = appendDialogNameField(
      nameHost,
      mac?.strings.nameLabel ?? t.cellStyleName,
      opts.initialName ?? t.cellStyleNormal,
    );

    const boxes = {} as Record<IncludeKey, HTMLInputElement>;
    if (mac) {
      const previewLabel = document.createElement('span');
      previewLabel.textContent = mac.strings.previewLabel;
      const preview = document.createElement('span');
      preview.className = 'fc-cellstyledlg__preview';
      preview.textContent = name.input.value;
      if (mac.preview?.fill) preview.style.background = mac.preview.fill;
      if (mac.preview?.color) preview.style.color = mac.preview.color;
      if (mac.preview?.bold) preview.style.fontWeight = '700';
      if (mac.preview?.italic) preview.style.fontStyle = 'italic';
      name.input.addEventListener('input', () => {
        preview.textContent = name.input.value;
      });
      nameHost.append(previewLabel, preview);

      const caption = document.createElement('div');
      caption.className = 'fc-cellstyledlg__caption';
      caption.textContent = mac.strings.caption;
      const table = document.createElement('div');
      table.className = 'fc-cellstyledlg__table';
      table.setAttribute('role', 'group');
      const head = document.createElement('div');
      head.className = 'fc-cellstyledlg__head';
      for (const label of [mac.strings.propertyColumn, mac.strings.formatColumn]) {
        const cell = document.createElement('span');
        cell.textContent = label;
        head.appendChild(cell);
      }
      table.appendChild(head);
      for (const key of ['number', 'alignment', 'font', 'border', 'fill', 'protection'] as const) {
        const row = document.createElement('div');
        row.className = 'fc-cellstyledlg__row';
        boxes[key] = appendCheckboxRow(row, mac.strings.properties[key], include[key], key);
        const summary = document.createElement('span');
        summary.className = 'fc-cellstyledlg__summary';
        summary.textContent = mac.summaries[key];
        row.appendChild(summary);
        table.appendChild(row);
      }
      shell.body.append(caption, table);
    } else {
      const legend = document.createElement('div');
      legend.className = 'fc-tb__dlg__label';
      legend.textContent = t.cellStyleIncludes;
      shell.body.appendChild(legend);
      boxes.number = appendCheckboxRow(
        shell.body,
        t.cellStyleIncludeNumber,
        include.number,
        'number',
      );
      boxes.alignment = appendCheckboxRow(
        shell.body,
        t.cellStyleIncludeAlignment,
        include.alignment,
        'alignment',
      );
      boxes.font = appendCheckboxRow(shell.body, t.cellStyleIncludeFont, include.font, 'font');
      boxes.border = appendCheckboxRow(
        shell.body,
        t.cellStyleIncludeBorder,
        include.border,
        'border',
      );
      boxes.fill = appendCheckboxRow(shell.body, t.cellStyleIncludeFill, include.fill, 'fill');
      boxes.protection = appendCheckboxRow(
        shell.body,
        t.cellStyleIncludeProtection,
        include.protection,
        'protection',
      );
    }
    const { number, alignment, font, border, fill, protection } = boxes;
    const errorRow = appendErrorRow(shell.body);
    if (mac) {
      const formatBtn = appendDialogButton(shell.footer, { label: mac.strings.formatButton });
      // Anchored on the left, away from Cancel / OK.
      formatBtn.style.marginRight = 'auto';
      formatBtn.addEventListener('click', () => mac.onFormat?.());
    }
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.strings.hyperlinkDialog.cancel,
      okLabel: opts.strings.hyperlinkDialog.ok,
    });

    const lifecycle = installDialogLifecycle<CellStyleDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): CellStyleDialogResult | null => {
      const label = name.valueOrError(errorRow, opts.strings.namedRangeDialog.errorEmptyName);
      if (label === null) {
        return null;
      }
      const result: CellStyleDialogResult = {
        name: label,
        include: {
          number: number.checked,
          alignment: alignment.checked,
          font: font.checked,
          border: border.checked,
          fill: fill.checked,
          protection: protection.checked,
        },
      };
      lifecycle.finish(result);
      return result;
    };
    okBtn.addEventListener('click', () => {
      onOk();
    });
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, name.focus);
  });
