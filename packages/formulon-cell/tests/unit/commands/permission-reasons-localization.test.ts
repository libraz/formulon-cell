import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  deny,
  localizedDecisionReason,
  PERMISSION_REASON_KEYS,
  PERMISSION_REASONS,
  type PermissionReasonKey,
  permissionReasonText,
} from '../../../src/commands/permission-reasons.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import { LOCALE_SURFACE_LOCALES } from '../../test-utils/locale-surface-manifest.js';

const sourceRoot = resolve(import.meta.dirname, '../../../src');

describe('built-in permission reasons', () => {
  it('lists a non-empty English reason for every key', () => {
    expect(PERMISSION_REASON_KEYS.length).toBeGreaterThan(0);
    for (const key of PERMISSION_REASON_KEYS) expect(PERMISSION_REASONS[key], key).not.toBe('');
  });

  it('localizes every key in all seven locales', () => {
    for (const locale of LOCALE_SURFACE_LOCALES) {
      for (const key of PERMISSION_REASON_KEYS) {
        const text = permissionReasonText(dictionaries[locale], key);
        expect(text, `${locale}:${key}`).toBeTruthy();
        expect(typeof text, `${locale}:${key}`).toBe('string');
      }
    }
  });

  it('builds every refusal through deny() so no producer bypasses the key list', () => {
    const producers = [
      'commands/built-in-command-policy.ts',
      'commands/interaction-authorizer.ts',
      'commands/interaction-effects.ts',
      'commands/interaction-policy.ts',
      'interact/context-menu-permissions.ts',
    ];
    for (const file of producers) {
      const lines = readFileSync(resolve(sourceRoot, file), 'utf8').split('\n');
      // The one literal refusal carries a thrown error's own message.
      const literal = lines.filter(
        (line) => /allowed: false,/.test(line) && !line.includes('error.message'),
      );
      expect(literal, file).toEqual([]);
    }
  });

  it('carries the key and the English reason on a built-in refusal', () => {
    const key: PermissionReasonKey = 'sheetProtected';
    expect(deny('protected', key)).toEqual({
      allowed: false,
      code: 'protected',
      reason: 'sheet is protected',
      reasonKey: 'sheetProtected',
    });
  });
});

describe('localizedDecisionReason', () => {
  it('projects a built-in reason into the dictionary language', () => {
    expect(localizedDecisionReason(deny('invalid', 'cellHoldsNote'), dictionaries.ja)).toBe(
      dictionaries.ja.permissionReasons.cellHoldsNote,
    );
    expect(localizedDecisionReason(deny('protected', 'sheetProtected'), dictionaries.de)).toBe(
      dictionaries.de.protection.protectedCellEdit,
    );
    expect(
      localizedDecisionReason(deny('protected', 'workbookStructureProtected'), dictionaries.fr),
    ).toBe(dictionaries.fr.ribbonMenu.workbookStructureProtectedBlocked);
  });

  it('shows a host-supplied reason unchanged', () => {
    expect(
      localizedDecisionReason(
        { code: 'operationDenied', reason: 'Ask your admin.' },
        dictionaries.ja,
      ),
    ).toBe('Ask your admin.');
  });

  it('shows a host reason unchanged even when it carries an unknown key', () => {
    const forged = { code: 'invalid', reason: 'Host text', reasonKey: 'nope' } as never;
    expect(localizedDecisionReason(forged, dictionaries.ja)).toBe('Host text');
  });

  it('words a refusal without a reason by its code, never the raw code', () => {
    for (const locale of LOCALE_SURFACE_LOCALES) {
      const text = localizedDecisionReason({ code: 'outOfBounds' }, dictionaries[locale]);
      expect(text, locale).toBe(dictionaries[locale].permissionReasons.codeOutOfBounds);
      expect(text, locale).not.toBe('outOfBounds');
    }
  });
});
