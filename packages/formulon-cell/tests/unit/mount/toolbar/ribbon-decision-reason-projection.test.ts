import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { dictionaries } from '../../../../src/i18n/strings.js';
import { attachRibbonInteractionPolicy } from '../../../../src/mount/toolbar-interaction-policy.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';

describe('ribbon projection of a refused built-in decision', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'ja' });
  });
  afterEach(() => sheet.dispose());

  const ribbon = (command: string) => {
    const host = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = command;
    host.appendChild(button);
    const projection = attachRibbonInteractionPolicy(host, {
      getInstance: () => sheet.instance,
      closeDynamicDropdowns: () => undefined,
    });
    return { button, projection };
  };

  it('writes the localized reason under the one reason key', () => {
    const { button, projection } = ribbon('sortAsc');
    try {
      setProtectedSheet(sheet.instance.store, 0, true);
      projection.project();
      const text = dictionaries.ja.protection.protectedCellEdit;
      expect(button.disabled).toBe(true);
      expect(button.dataset.ribbonDisabledReason).toBe(text);
      expect(button.getAttribute('aria-description')).toBe(text);
      expect(button.title).toBe(text);
      expect(button.dataset.ribbonDecisionReason).toBeUndefined();

      setProtectedSheet(sheet.instance.store, 0, false);
      projection.project();
      expect(button.disabled).toBe(false);
      expect(button.dataset.ribbonDisabledReason).toBeUndefined();
    } finally {
      projection.detach();
    }
  });

  it('shows a host policy reason unchanged', () => {
    sheet.instance.setPolicy({
      defaultOperation: 'allow',
      restrict: ({ commandId }) =>
        commandId === 'sortAsc'
          ? { allowed: false, code: 'operationDenied', reason: 'Sorting is locked by IT.' }
          : true,
    });
    const { button, projection } = ribbon('sortAsc');
    try {
      projection.project();
      expect(button.disabled).toBe(true);
      expect(button.dataset.ribbonDisabledReason).toBe('Sorting is locked by IT.');
      expect(button.title).toBe('Sorting is locked by IT.');
    } finally {
      projection.detach();
    }
  });

  it('localizes the generic refusal of a host policy that gives no reason', () => {
    sheet.instance.setPolicy({ operations: { sort: false } });
    const { button, projection } = ribbon('sortAsc');
    try {
      projection.project();
      expect(button.disabled).toBe(true);
      expect(button.dataset.ribbonDisabledReason).toBe(
        dictionaries.ja.permissionReasons.operationDenied,
      );
    } finally {
      projection.detach();
    }
  });
});
