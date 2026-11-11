import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { attachRibbonInteractionPolicy } from '../../../src/mount/toolbar-interaction-policy.js';
import { projectDisabledState } from '../../../src/toolbar/menu-a11y.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('disabled-state projection and the element base title', () => {
  it('restores the title the element had before it was disabled', () => {
    const button = document.createElement('button');
    button.title = 'Plot Right-to-Left';
    projectDisabledState(button, true, 'Not available');
    expect(button.title).toBe('Not available');
    projectDisabledState(button, false, null);
    expect(button.title).toBe('Plot Right-to-Left');
  });

  it('keeps the first base title across a change of reason', () => {
    const button = document.createElement('button');
    button.title = 'Style';
    projectDisabledState(button, true, 'First reason');
    projectDisabledState(button, true, 'Second reason');
    expect(button.title).toBe('Second reason');
    projectDisabledState(button, false, null);
    expect(button.title).toBe('Style');
  });

  it('restores an empty title for an element that had none', () => {
    const button = document.createElement('button');
    projectDisabledState(button, true, 'Not available');
    projectDisabledState(button, false, null);
    expect(button.title).toBe('');
  });

  it('leaves the title of an element that was never disabled', () => {
    const button = document.createElement('button');
    button.title = 'Bold';
    projectDisabledState(button, false, null);
    expect(button.title).toBe('Bold');
  });

  it('lets an explicit title prefix win over the remembered title', () => {
    const button = document.createElement('button');
    button.title = 'Ignored';
    projectDisabledState(button, true, 'Not available', { titlePrefix: 'Sum' });
    expect(button.title).toBe('Sum\nNot available');
    projectDisabledState(button, false, null, { titlePrefix: 'Sum' });
    expect(button.title).toBe('Sum');
  });
});

describe('ribbon policy projection and the base title', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
  });
  afterEach(() => sheet.dispose());

  it('restores the button title when protection lifts', () => {
    const host = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = 'sortAsc';
    button.title = 'Sort A to Z';
    host.appendChild(button);
    const projection = attachRibbonInteractionPolicy(host, {
      getInstance: () => sheet.instance,
      closeDynamicDropdowns: () => undefined,
    });
    try {
      setProtectedSheet(sheet.instance.store, 0, true);
      projection.project();
      expect(button.disabled).toBe(true);
      expect(button.title).not.toBe('Sort A to Z');
      setProtectedSheet(sheet.instance.store, 0, false);
      projection.project();
      expect(button.disabled).toBe(false);
      expect(button.title).toBe('Sort A to Z');
    } finally {
      projection.detach();
    }
  });
});
