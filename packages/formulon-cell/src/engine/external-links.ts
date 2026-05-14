export type ExternalLinkKind = 'unknown' | 'externalBook' | 'ole' | 'dde';

export const externalLinkKindLabel = (kind: number): ExternalLinkKind => {
  switch (kind) {
    case 1:
      return 'externalBook';
    case 2:
      return 'ole';
    case 3:
      return 'dde';
    default:
      return 'unknown';
  }
};
