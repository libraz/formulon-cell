import type { SpreadsheetInstance } from '../../api/spreadsheet/instance.js';
import { workbookHandleOf } from '../../api/workbook.js';
import {
  type WorkbookSaveFormat,
  type WorkbookSaveLoss,
  type WorkbookSaveOutcome,
  writeWorkbook,
  type XlsbUnwritable,
  xlsbUnwritableOf,
} from '../../commands/workbook-save.js';
import type { FileLossId } from '../../engine/file-diagnostics.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { Strings } from '../../i18n/strings.js';
import { createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
  appendDialogMessage,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

const formatLabel = (strings: Strings, format: WorkbookSaveFormat): string =>
  format === 'xlsb' ? strings.workbookSave.formatXlsb : strings.workbookSave.formatXlsx;

const lossLabel = (strings: Strings, loss: WorkbookSaveLoss): string => {
  const t = strings.workbookSave;
  switch (loss.id) {
    case 'notes':
    case 'threadedComments':
    case 'autoFilters':
    case 'pivotTables':
    case 'slicers':
    case 'drawings':
    case 'tables':
    case 'sparklines':
    case 'protectedRanges':
    case 'otherSettings':
      return t[loss.id];
    default:
      return strings.workbookObjects.lossLabels[loss.id as FileLossId];
  }
};

const lossText = (strings: Strings, loss: WorkbookSaveLoss): string => {
  const label = lossLabel(strings, loss);
  return loss.count === undefined ? label : `${label} (${loss.count})`;
};

const appendLossList = (body: HTMLElement, items: readonly string[]): void => {
  const list = body.ownerDocument.createElement('ul');
  list.className = 'fc-tb__dlg__loss-list';
  for (const text of items) {
    const item = body.ownerDocument.createElement('li');
    item.textContent = text;
    list.appendChild(item);
  }
  body.appendChild(list);
};

/** An alert dialog with a filled body and Cancel / OK; resolves true on OK. */
const confirmWithBody = (
  title: string,
  okLabel: string,
  cancelLabel: string,
  fillBody: (body: HTMLElement) => void,
  host?: HTMLElement,
): Promise<boolean> =>
  new Promise<boolean>((resolve) => {
    const shell = createDialogShell({ host, title, role: 'alertdialog', bodyVariant: 'app' });
    fillBody(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, { cancelLabel, okLabel });
    const lifecycle = installDialogLifecycle<boolean>({
      shell,
      resolve,
      onCancel: () => false,
    });
    okBtn.addEventListener('click', () => lifecycle.finish(true));
    cancelBtn.addEventListener('click', () => lifecycle.finish(false));
    mountDialog(shell, cancelBtn);
  });

/** Warn that `format` leaves `losses` out. Resolves true to save anyway. */
export const confirmWorkbookSaveLosses = (
  strings: Strings,
  format: WorkbookSaveFormat,
  losses: readonly WorkbookSaveLoss[],
  host?: HTMLElement,
): Promise<boolean> => {
  const t = strings.workbookSave;
  const type = formatLabel(strings, format);
  const certain = losses.filter((loss) => loss.certain);
  const possible = losses.filter((loss) => !loss.certain);
  return confirmWithBody(
    t.lossTitle,
    t.saveAnyway,
    t.cancel,
    (body) => {
      if (certain.length > 0) {
        appendDialogMessage(body, t.lossMessage.replace('{format}', type));
        appendLossList(
          body,
          certain.map((loss) => lossText(strings, loss)),
        );
      }
      if (possible.length > 0) {
        appendDialogMessage(body, t.mayLoseMessage.replace('{format}', type));
        appendLossList(
          body,
          possible.map((loss) => lossText(strings, loss)),
        );
      }
    },
    host,
  );
};

/** Tell the user the `.xlsb` writer cannot hold `feature`, on sheet `sheetName` when it is
 *  sheet-scoped. Resolves true to save as `.xlsx` instead. */
const offerXlsxForUnwritable = (
  strings: Strings,
  feature: XlsbUnwritable,
  sheetName: string | null,
  host?: HTMLElement,
): Promise<boolean> => {
  const t = strings.workbookSave;
  return confirmWithBody(
    t.unwritableTitle,
    t.saveAsFormat.replace('{format}', formatLabel(strings, 'xlsx')),
    t.cancel,
    (body) => {
      appendDialogMessage(body, t.unwritableMessage);
      appendLossList(body, [
        sheetName === null ? t[feature.id] : `${t[feature.id]} (${sheetName})`,
      ]);
    },
    host,
  );
};

/** File types `wb` can be written as, in menu order. */
export const workbookSaveFormats = (wb: WorkbookHandle): WorkbookSaveFormat[] =>
  wb.capabilities.xlsbSave ? ['xlsx', 'xlsb'] : ['xlsx'];

/** Save As: pick the file type among those `wb` can be written as. Resolves
 *  null when cancelled. */
export const chooseWorkbookSaveFormat = (
  wb: WorkbookHandle,
  strings: Strings,
  host?: HTMLElement,
): Promise<WorkbookSaveFormat | null> =>
  new Promise<WorkbookSaveFormat | null>((resolve) => {
    const t = strings.workbookSave;
    const formats = workbookSaveFormats(wb);
    const shell = createDialogShell({ host, title: t.saveAsTitle, bodyVariant: 'app' });
    const row = shell.body.ownerDocument.createElement('label');
    row.className = 'fc-fmtdlg__row';
    const label = shell.body.ownerDocument.createElement('span');
    label.textContent = t.fileFormat;
    const select = createDialogSelect(
      formats.map((format) => ({ value: format, label: formatLabel(strings, format) })),
      'xlsx',
      { ariaLabel: t.fileFormat },
    );
    select.dataset.fcSaveFormat = '';
    row.append(label, select);
    shell.body.appendChild(row);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: t.cancel,
      okLabel: t.save,
    });
    const chosen = (): WorkbookSaveFormat =>
      formats.find((format) => format === select.value) ?? 'xlsx';
    const lifecycle = installDialogLifecycle<WorkbookSaveFormat | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => lifecycle.finish(chosen()),
      buttons: { okBtn, cancelBtn },
    });
    mountDialog(shell, select);
  });

/** Options of {@link saveWithPrompts}. */
export interface SaveWithPromptsOptions {
  /** The file type to write. Omitted, the Save As dialog asks for it. */
  format?: WorkbookSaveFormat;
  /** Element the dialogs mount in; the spreadsheet's own host by default. */
  host?: HTMLElement;
}

/** Write the spreadsheet's workbook and ask before handing over a file that
 *  leaves content out. Without `format` the Save As dialog picks the file type
 *  first. An `.xlsb` save the writer refuses offers `.xlsx` instead. Resolves
 *  the bytes to deliver, or null when the user cancels. */
export async function saveWithPrompts(
  spreadsheet: SpreadsheetInstance,
  options: SaveWithPromptsOptions = {},
): Promise<WorkbookSaveOutcome | null> {
  const wb = workbookHandleOf(spreadsheet.workbook);
  const strings = spreadsheet.i18n.strings;
  const host = options.host ?? spreadsheet.host;
  let format = options.format ?? (await chooseWorkbookSaveFormat(wb, strings, host));
  if (format === null) return null;
  let outcome: WorkbookSaveOutcome;
  try {
    outcome = writeWorkbook(wb, format);
  } catch (error) {
    const refused = format === 'xlsb' ? xlsbUnwritableOf(error) : null;
    if (!refused) throw error;
    const sheet = refused.sheet === null ? null : wb.sheetName(refused.sheet);
    if (!(await offerXlsxForUnwritable(strings, refused, sheet, host))) return null;
    format = 'xlsx';
    outcome = writeWorkbook(wb, format);
  }
  if (outcome.losses.length === 0) return outcome;
  const proceed = await confirmWorkbookSaveLosses(strings, format, outcome.losses, host);
  return proceed ? outcome : null;
}
