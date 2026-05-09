import { describe, expect, it } from 'vitest';
import {
  type ExternalLinkRecord,
  listExternalLinks,
  summarizeExternalLinks,
} from '../../../src/commands/external-links.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const fakeWb = (links: readonly ExternalLinkRecord[]): WorkbookHandle =>
  ({ getExternalLinks: () => links }) as unknown as WorkbookHandle;

describe('external link commands', () => {
  it('lists external links from the workbook handle', () => {
    const links: ExternalLinkRecord[] = [
      {
        index: 1,
        relId: 'rId1',
        partPath: 'xl/externalLinks/externalLink1.xml',
        target: 'file:///book.xlsx',
        targetExternal: true,
        kind: 'externalBook',
      },
    ];

    expect(listExternalLinks(fakeWb(links))).toEqual(links);
    expect(listExternalLinks(null)).toEqual([]);
  });

  it('summarizes count, external targets, and link kinds', () => {
    const summary = summarizeExternalLinks(
      fakeWb([
        {
          index: 1,
          relId: 'rId1',
          partPath: 'xl/externalLinks/externalLink1.xml',
          target: 'file:///book.xlsx',
          targetExternal: true,
          kind: 'externalBook',
        },
        {
          index: 2,
          relId: 'rId2',
          partPath: 'xl/externalLinks/externalLink2.xml',
          target: '',
          targetExternal: false,
          kind: 'unknown',
        },
      ]),
    );

    expect(summary.count).toBe(2);
    expect(summary.externalCount).toBe(1);
    expect(summary.byKind).toEqual({ unknown: 1, externalBook: 1, ole: 0, dde: 0 });
  });
});
