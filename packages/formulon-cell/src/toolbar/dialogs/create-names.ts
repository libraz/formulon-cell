import { projectDisabledState } from '../menu-a11y.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

/** The selection's edges whose labels name the rest of each row or column. */
export interface CreateNamesEdges {
  top: boolean;
  left: boolean;
  bottom: boolean;
  right: boolean;
}

export interface CreateNamesDialogOptions {
  /** The spreadsheet host; anchors the dialog and picks its platform layout. */
  host?: HTMLElement;
  labels: {
    title: string;
    createIn: string;
    top: string;
    left: string;
    bottom: string;
    right: string;
    ok: string;
    cancel: string;
  };
  /** The edges checked on open. */
  initial: CreateNamesEdges;
}

const EDGES = ['top', 'left', 'bottom', 'right'] as const;

/** The Create Names dialog: create names in any combination of the top row, left column, bottom
 *  row and right column. Resolves with the checked edges on OK, and `null` on Cancel. OK stays
 *  greyed while no edge is checked. */
export const showCreateNamesDialog = (
  opts: CreateNamesDialogOptions,
): Promise<CreateNamesEdges | null> =>
  new Promise<CreateNamesEdges | null>((resolve) => {
    const t = opts.labels;
    const shell = createDialogShell({ host: opts.host, title: t.title, bodyVariant: 'app' });
    shell.panel.classList.add('fc-createnames');
    const doc = shell.body.ownerDocument;
    const group = doc.createElement('fieldset');
    group.className = 'fc-createnames__edges';
    const legend = doc.createElement('legend');
    legend.textContent = t.createIn;
    group.appendChild(legend);
    const boxes = {} as Record<(typeof EDGES)[number], HTMLInputElement>;
    for (const edge of EDGES) {
      const row = doc.createElement('label');
      row.className = 'fc-createnames__edge';
      const box = doc.createElement('input');
      box.type = 'checkbox';
      box.dataset.fcEdge = edge;
      box.checked = opts.initial[edge];
      row.append(box, doc.createTextNode(` ${t[edge]}`));
      group.appendChild(row);
      boxes[edge] = box;
    }
    shell.body.appendChild(group);

    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: t.cancel,
      okLabel: t.ok,
    });
    const edges = (): CreateNamesEdges => ({
      top: boxes.top.checked,
      left: boxes.left.checked,
      bottom: boxes.bottom.checked,
      right: boxes.right.checked,
    });
    const anyChecked = (): boolean => EDGES.some((edge) => boxes[edge].checked);
    const sync = (): void => projectDisabledState(okBtn, !anyChecked(), null);
    for (const edge of EDGES) boxes[edge].addEventListener('change', sync);
    sync();

    const lifecycle = installDialogLifecycle<CreateNamesEdges | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => {
        if (anyChecked()) lifecycle.finish(edges());
      },
    });
    okBtn.addEventListener('click', () => lifecycle.finish(edges()));
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, boxes.top, { immediate: true });
  });
