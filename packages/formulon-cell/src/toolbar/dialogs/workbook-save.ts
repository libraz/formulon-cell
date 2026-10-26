import {
  type WorkbookSaveFormat,
  type WorkbookSaveLoss,
  type WorkbookSaveOutcome,
  writeWorkbook,
} from '../../commands/workbook-save.js';
import type { FileLossId } from '../../engine/file-diagnostics.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { Strings } from '../../i18n/strings.js';
import { createDialogSelect } from './form-controls.js';
import {
  appendDialogActions,
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

const appendMessage = (body: HTMLElement, text: string): void => {
  const msg = body.ownerDocument.createElement('p');
  msg.className = 'fc-tb__dlg__message';
  msg.textContent = text;
  body.appendChild(msg);
};

const appendLossList = (
  body: HTMLElement,
  strings: Strings,
  losses: readonly WorkbookSaveLoss[],
) => {
  const list = body.ownerDocument.createElement('ul');
  list.className = 'fc-tb__dlg__loss-list';
  for (const loss of losses) {
    const item = body.ownerDocument.createElement('li');
    const label = lossLabel(strings, loss);
    item.textContent = loss.count === undefined ? label : `${label} (${loss.count})`;
    list.appendChild(item);
  }
  body.appendChild(list);
};

/** Warn that `format` leaves `losses` out. Resolves true to save anyway. */
export const confirmWorkbookSaveLosses = (
  strings: Strings,
  format: WorkbookSaveFormat,
  losses: readonly WorkbookSaveLoss[],
  host?: HTMLElement,
): Promise<boolean> =>
  new Promise<boolean>((resolve) => {
    const t = strings.workbookSave;
    const shell = createDialogShell({
      host,
      title: t.lossTitle,
      role: 'alertdialog',
      bodyVariant: 'app',
    });
    const type = formatLabel(strings, format);
    const certain = losses.filter((loss) => loss.certain);
    const possible = losses.filter((loss) => !loss.certain);
    if (certain.length > 0) {
      appendMessage(shell.body, t.lossMessage.replace('{format}', type));
      appendLossList(shell.body, strings, certain);
    }
    if (possible.length > 0) {
      appendMessage(shell.body, t.mayLoseMessage.replace('{format}', type));
      appendLossList(shell.body, strings, possible);
    }
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: t.cancel,
      okLabel: t.saveAnyway,
    });
    const lifecycle = installDialogLifecycle<boolean>({
      shell,
      resolve,
      onCancel: () => false,
    });
    okBtn.addEventListener('click', () => lifecycle.finish(true));
    cancelBtn.addEventListener('click', () => lifecycle.finish(false));
    mountDialog(shell, cancelBtn);
  });

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
    });
    okBtn.addEventListener('click', () => lifecycle.finish(chosen()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, select);
  });

export interface SaveWorkbookWithPromptsOptions {
  wb: WorkbookHandle;
  strings: Strings;
  format: WorkbookSaveFormat;
  host?: HTMLElement;
}

/** Write the workbook and ask before handing over a file that leaves content
 *  out. Resolves the bytes to deliver, or null when the user cancels. */
export async function saveWorkbookWithPrompts(
  opts: SaveWorkbookWithPromptsOptions,
): Promise<WorkbookSaveOutcome | null> {
  const outcome = writeWorkbook(opts.wb, opts.format);
  if (outcome.losses.length === 0) return outcome;
  const proceed = await confirmWorkbookSaveLosses(
    opts.strings,
    opts.format,
    outcome.losses,
    opts.host,
  );
  return proceed ? outcome : null;
}
