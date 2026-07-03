import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachExternalLinksDialog } from '../../../src/interact/external-links-dialog.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
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
    const dialog = document.querySelector<HTMLElement>('.fc-extlinkdlg');
    expect(dialog).not.toBeNull();
    expect(dialog?.hidden).toBe(false);
    const empty = document.querySelector<HTMLElement>('.fc-extlinkdlg__empty');
    expect(empty?.hidden).toBe(false);
    expect(document.querySelector('.fc-extlinkdlg__table')).toBeNull();
    const actions = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-extlinkdlg__action'),
    );
    expect(actions.map((button) => button.textContent)).toEqual([
      defaultStrings.externalLinksDialog.updateValues,
      defaultStrings.externalLinksDialog.changeSource,
      defaultStrings.externalLinksDialog.openSource,
      defaultStrings.externalLinksDialog.breakLink,
      defaultStrings.externalLinksDialog.checkStatus,
      defaultStrings.externalLinksDialog.startupPrompt,
    ]);
    for (const action of actions) {
      expect(action.disabled).toBe(true);
      expect(action.getAttribute('aria-disabled')).toBe('true');
      expect(action.dataset.disabledReason).toBe(
        defaultStrings.externalLinksDialog.noSelectionActionReason,
      );
      expect(action.getAttribute('aria-description')).toBe(
        defaultStrings.externalLinksDialog.noSelectionActionReason,
      );
    }
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
    const rows = document.querySelectorAll<HTMLTableRowElement>('.fc-extlinkdlg__table tbody tr');
    expect(rows.length).toBe(2);
    expect(rows[0]?.textContent).toContain('1');
    expect(rows[0]?.textContent).toContain(defaultStrings.externalLinksDialog.kindExternalBook);
    expect(rows[0]?.textContent).toContain('book2.xlsx');
    expect(rows[0]?.tabIndex).toBe(0);
    expect(rows[0]?.getAttribute('aria-selected')).toBe('true');
    expect(rows[1]?.tabIndex).toBe(-1);
    expect(rows[1]?.textContent).toContain(defaultStrings.externalLinksDialog.kindUnknown);
    // Empty target renders as a dash placeholder.
    expect(rows[1]?.textContent).toContain('—');
    const actions = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-extlinkdlg__action'),
    );
    expect(actions).toHaveLength(6);
    for (const action of actions) {
      expect(action.disabled).toBe(true);
      expect(action.dataset.disabledReason).toBe(
        defaultStrings.externalLinksDialog.readOnlyActionReason,
      );
    }
    handle.detach();
  });

  it('selects rows by click while keeping Edit Links actions read-only', () => {
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
    const rows = Array.from(
      document.querySelectorAll<HTMLTableRowElement>('.fc-extlinkdlg__table tbody tr'),
    );
    rows[1]?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(rows[0]?.getAttribute('aria-selected')).toBe('false');
    expect(rows[1]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows[1]);
    expect(
      document.querySelector<HTMLButtonElement>('[data-external-link-action="breakLink"]')?.dataset
        .disabledReason,
    ).toBe(defaultStrings.externalLinksDialog.readOnlyActionReason);
    handle.detach();
  });

  it('supports Excel-style row navigation keys', () => {
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
    const rows = (): HTMLTableRowElement[] =>
      Array.from(document.querySelectorAll<HTMLTableRowElement>('.fc-extlinkdlg__table tbody tr'));

    rows()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(rows()[1]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows()[1]);

    rows()[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(rows()[0]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows()[0]);

    rows()[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(rows()[1]?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('Escape closes the dialog', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
    });
    handle.open();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const dialog = document.querySelector<HTMLElement>('.fc-extlinkdlg');
    expect(dialog?.hidden).toBe(true);
    handle.detach();
  });

  it('clicking the overlay backdrop closes the dialog', () => {
    const handle = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
    });
    handle.open();
    const dialog = document.querySelector<HTMLElement>('.fc-extlinkdlg');
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
    expect(document.querySelector('.fc-extlinkdlg')).toBeNull();
  });

  it('keeps action buttons on the shared dialog button helper', () => {
    const source = readFileSync(join(root, 'src/interact/external-links-dialog.ts'), 'utf8');
    expect(source).toContain('appendExternalLinkActionButton(actionBar');
    expect(source).toContain('appendDialogButton(parent');
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps External Links on compact desktop table geometry', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/dialogs/external-links.css'), 'utf8');

    expect(css).toMatch(/\.fc-extlinkdlg__panel\s*\{[\s\S]*?border-radius: 2px;/);
    expect(css).toMatch(
      /\.fc-extlinkdlg__table tbody tr\[aria-selected="true"\]\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(css).not.toContain('background: var(--fc-accent-soft');
  });
});
