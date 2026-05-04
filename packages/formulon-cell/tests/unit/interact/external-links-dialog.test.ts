import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachExternalLinksDialog } from '../../../src/interact/external-links-dialog.js';

type LinkRecord = ReturnType<WorkbookHandle['getExternalLinks']>[number];

const fakeWb = (links: readonly LinkRecord[]): WorkbookHandle =>
  ({
    getExternalLinks: () => links,
  }) as unknown as WorkbookHandle;

describe('attachExternalLinksDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('renders an empty-state message when the workbook has no externals', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
    });
    handle.open();
    const dialog = host.querySelector<HTMLElement>('.fc-extlinkdlg');
    expect(dialog).not.toBeNull();
    expect(dialog?.hidden).toBe(false);
    const empty = host.querySelector<HTMLElement>('.fc-extlinkdlg__empty');
    expect(empty?.hidden).toBe(false);
    expect(host.querySelector('.fc-extlinkdlg__table')).toBeNull();
    handle.detach();
  });

  it('renders one row per external link with index/kind/target/part columns', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () =>
        fakeWb([
          {
            index: 1,
            relId: 'rId3',
            partPath: 'xl/externalLinks/externalLink1.xml',
            target: 'file:///fixtures/book2.xlsx',
            targetExternal: true,
            kind: 'externalBook',
          },
          {
            index: 2,
            relId: 'rId7',
            partPath: 'xl/externalLinks/externalLink2.xml',
            target: '',
            targetExternal: false,
            kind: 'unknown',
          },
        ]),
    });
    handle.open();
    const rows = host.querySelectorAll<HTMLTableRowElement>('.fc-extlinkdlg__table tbody tr');
    expect(rows.length).toBe(2);
    expect(rows[0]?.textContent).toContain('1');
    expect(rows[0]?.textContent).toContain('externalBook');
    expect(rows[0]?.textContent).toContain('book2.xlsx');
    // Empty target renders as a dash placeholder.
    expect(rows[1]?.textContent).toContain('—');
    handle.detach();
  });

  it('Escape closes the dialog', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
    });
    handle.open();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const dialog = host.querySelector<HTMLElement>('.fc-extlinkdlg');
    expect(dialog?.hidden).toBe(true);
    handle.detach();
  });

  it('clicking the overlay backdrop closes the dialog', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
    });
    handle.open();
    const dialog = host.querySelector<HTMLElement>('.fc-extlinkdlg');
    dialog?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    expect(dialog?.hidden).toBe(true);
    handle.detach();
  });

  it('detach removes the dialog node from the DOM', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
    });
    handle.detach();
    expect(host.querySelector('.fc-extlinkdlg')).toBeNull();
  });
});
