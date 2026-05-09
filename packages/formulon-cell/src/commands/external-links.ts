import type { WorkbookHandle } from '../engine/workbook-handle.js';

export type ExternalLinkKind = 'unknown' | 'externalBook' | 'ole' | 'dde';

export interface ExternalLinkRecord {
  index: number;
  relId: string;
  partPath: string;
  target: string;
  targetExternal: boolean;
  kind: ExternalLinkKind;
}

export interface ExternalLinksSummary {
  count: number;
  externalCount: number;
  byKind: Record<ExternalLinkKind, number>;
  links: readonly ExternalLinkRecord[];
}

export function listExternalLinks(
  workbook: WorkbookHandle | null | undefined,
): ExternalLinkRecord[] {
  return [...(workbook?.getExternalLinks() ?? [])];
}

export function summarizeExternalLinks(
  workbook: WorkbookHandle | null | undefined,
): ExternalLinksSummary {
  const links = listExternalLinks(workbook);
  const byKind: Record<ExternalLinkKind, number> = {
    unknown: 0,
    externalBook: 0,
    ole: 0,
    dde: 0,
  };
  let externalCount = 0;
  for (const link of links) {
    byKind[link.kind] += 1;
    if (link.targetExternal) externalCount += 1;
  }
  return { count: links.length, externalCount, byKind, links };
}
