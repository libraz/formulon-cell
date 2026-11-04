import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../../../src/i18n/strings.js';
import { attachExternalLinksDialog } from '../../../src/interact/external-links-dialog.js';

type LinkRecord = ReturnType<WorkbookHandle['getExternalLinks']>[number];
const fakeWb = (links: readonly LinkRecord[]): WorkbookHandle =>
  ({ getExternalLinks: () => links }) as unknown as WorkbookHandle;

// Excel for Mac (ja): the links are a docked "ブックのリンク" pane, not a modal.
const strings = {
  ...defaultStrings,
  macExternalLinksPane: {
    title: 'ブックのリンク',
    close: 'ブックのリンク を閉じる',
    empty: '外部ブックへのリンクはありません。',
    help: 'ブック リンクの作成方法の詳細',
  },
} as unknown as Strings;

describe('Mac external links pane', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.replaceChildren();
  });

  it('docks a pane in the host instead of opening the modal', () => {
    const handle = attachExternalLinksDialog({ host, getWb: () => fakeWb([]), strings });
    handle.open();
    const pane = host.querySelector<HTMLElement>('.fc-extlinkpane');
    expect(pane?.hidden).toBe(false);
    expect(pane?.getAttribute('aria-label')).toBe('ブックのリンク');
    expect(host.querySelector('.fc-extlinkpane__title')?.textContent).toBe('ブックのリンク');
    expect(host.querySelector('.fc-extlinkpane__close')?.getAttribute('aria-label')).toBe(
      'ブックのリンク を閉じる',
    );
    expect(host.querySelector('.fc-extlinkpane__empty')?.textContent).toBe(
      '外部ブックへのリンクはありません。',
    );
    expect(host.querySelector('.fc-extlinkpane__help')?.textContent).toBe(
      'ブック リンクの作成方法の詳細',
    );
    expect(document.querySelector('.fc-extlinkdlg:not([hidden])')).toBeNull();
    handle.detach();
  });

  it('draws the close button as a filled disc with an X cut out of it', () => {
    const handle = attachExternalLinksDialog({ host, getWb: () => fakeWb([]), strings });
    handle.open();
    const svg = host.querySelector('.fc-extlinkpane__close svg');
    expect(svg?.querySelectorAll('path')).toHaveLength(2);
    expect(svg?.querySelector<SVGPathElement>('path[fill="none"]')?.style.stroke).toContain(
      '--fc-tb-icon-paper',
    );
  });

  it('closes from the round X button and reopens', () => {
    const handle = attachExternalLinksDialog({ host, getWb: () => fakeWb([]), strings });
    handle.open();
    host.querySelector<HTMLButtonElement>('.fc-extlinkpane__close')?.click();
    expect(host.querySelector<HTMLElement>('.fc-extlinkpane')?.hidden).toBe(true);
    handle.open();
    expect(host.querySelector<HTMLElement>('.fc-extlinkpane')?.hidden).toBe(false);
    handle.detach();
  });

  it('lists the workbook links when there are any', () => {
    const links = [
      {
        index: 1,
        relId: 'rId1',
        partPath: 'xl/externalLinks/externalLink1.xml',
        target: 'file:///other.xlsx',
        targetExternal: true,
        kind: 'externalBook',
      },
    ] as unknown as readonly LinkRecord[];
    const handle = attachExternalLinksDialog({ host, getWb: () => fakeWb(links), strings });
    handle.open();
    expect(host.querySelector<HTMLElement>('.fc-extlinkpane__empty')?.hidden).toBe(true);
    const rows = Array.from(host.querySelectorAll('.fc-extlinkpane__link'), (r) => r.textContent);
    expect(rows).toEqual([expect.stringContaining('file:///other.xlsx')]);
    handle.detach();
  });

  it('keeps the modal when the host is not Mac', () => {
    delete host.dataset.fcPlatform;
    const plain = attachExternalLinksDialog({
      host,
      getWb: () => fakeWb([]),
      strings: defaultStrings,
    });
    plain.open();
    expect(host.querySelector('.fc-extlinkpane')).toBeNull();
    expect(document.querySelector('.fc-extlinkdlg')).not.toBeNull();
    plain.detach();
  });

  it('docks into the host task-pane dock and hides it again on close', () => {
    const dock = document.createElement('div');
    dock.className = 'fc-host__taskpane-dock';
    dock.hidden = true;
    const other = document.createElement('div');
    other.hidden = true;
    dock.appendChild(other);
    host.appendChild(dock);
    const handle = attachExternalLinksDialog({ host, getWb: () => fakeWb([]), strings });
    handle.open();
    const pane = host.querySelector<HTMLElement>('.fc-extlinkpane');
    expect(pane?.parentElement).toBe(dock);
    expect(pane?.style.top).toBe('');
    expect(dock.hidden).toBe(false);
    other.hidden = false;
    handle.close();
    expect(dock.hidden, 'another visible pane keeps the dock open').toBe(false);
    other.hidden = true;
    handle.open();
    host.querySelector<HTMLButtonElement>('.fc-extlinkpane__close')?.click();
    expect(dock.hidden).toBe(true);
    handle.detach();
  });
});
