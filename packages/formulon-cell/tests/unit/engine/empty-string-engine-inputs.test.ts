import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const PERSON_ID = '{11111111-1111-4111-8111-111111111111}';
const ROOT_ID = '{22222222-2222-4222-8222-222222222222}';

// The engine refuses an empty string for these inputs; the handle maps it to the absent value.
describe.skipIf(!canLoadWasm())('empty strings handed to the engine', () => {
  it('treats empty workbook protection hash fields as absent', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const written = wb.setWorkbookProtection({
        lockStructure: true,
        workbookLegacyPasswordHash: '',
        workbookAlgorithmName: 'SHA-512',
        workbookHashValue: 'aGFzaA==',
        workbookSaltValue: '',
        workbookSpinCount: 10,
        revisionsAlgorithmName: '',
        revisionsHashValue: '',
        revisionsSaltValue: '',
        revisionsLegacyPasswordHash: '',
      });
      expect(written).toBe(true);
      expect(wb.getWorkbookProtection()).toMatchObject({
        lockStructure: true,
        workbookAlgorithmName: 'SHA-512',
        workbookSaltValue: null,
        workbookLegacyPasswordHash: null,
        revisionsAlgorithmName: null,
      });
    } finally {
      wb.dispose();
    }
  });

  it('treats empty protected range hash fields as absent when adding', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const added = wb.addProtectedRange(0, {
        name: 'Editable',
        ranges: 'A1:B2',
        legacyPasswordHash: '',
        algorithmName: '',
        hashValue: '',
        saltValue: '',
        securityDescriptor: '',
      });
      expect(added).toBe(true);
      expect(wb.getProtectedRanges(0)[0]).toMatchObject({
        name: 'Editable',
        legacyPasswordHash: null,
        algorithmName: null,
        hashValue: null,
        saltValue: null,
        securityDescriptor: null,
      });
    } finally {
      wb.dispose();
    }
  });

  it('treats empty protected range hash fields as absent when updating', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(
        wb.addProtectedRange(0, {
          name: 'Editable',
          ranges: 'A1:B2',
          legacyPasswordHash: 'ABCD',
          securityDescriptor: 'O:BAG:BAD:(A;;GA;;;WD)',
        }),
      ).toBe(true);
      const updated = wb.updateProtectedRange(0, 0, {
        name: 'Editable',
        ranges: 'A1:B2',
        legacyPasswordHash: '',
        algorithmName: '',
        hashValue: '',
        saltValue: '',
        securityDescriptor: '',
      });
      expect(updated).toBe(true);
      expect(wb.getProtectedRanges(0)[0]).toMatchObject({
        legacyPasswordHash: null,
        securityDescriptor: null,
      });
    } finally {
      wb.dispose();
    }
  });

  it('opens a new thread when the reply target is empty', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.addPerson({ id: PERSON_ID, displayName: 'Ann' })).toBe(true);
      const added = wb.addThreadedComment(0, {
        id: ROOT_ID,
        cell: { row: 0, col: 0 },
        personId: PERSON_ID,
        text: 'hello',
        parentId: '',
      });
      expect(added).toBe(true);
      expect(wb.getThreadedComments(0)[0]).toMatchObject({ id: ROOT_ID, parentId: null });
    } finally {
      wb.dispose();
    }
  });

  it('generates a mention id when the given one is empty', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.addPerson({ id: PERSON_ID, displayName: 'Ann' })).toBe(true);
      const added = wb.addThreadedComment(0, {
        id: ROOT_ID,
        cell: { row: 0, col: 0 },
        personId: PERSON_ID,
        text: '@Ann hello',
        mentions: [{ personId: PERSON_ID, mentionId: '', start: 0, end: 4 }],
      });
      expect(added).toBe(true);
      expect(wb.getThreadedComments(0)[0]?.mentions).toHaveLength(1);
    } finally {
      wb.dispose();
    }
  });
});
