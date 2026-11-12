import { afterEach, describe, expect, it } from 'vitest';
import { setComment } from '../../../src/commands/comment.js';
import { en } from '../../../src/i18n/strings.js';
import { mutators } from '../../../src/store/store.js';
import { projectRibbonActiveState } from '../../../src/toolbar/ribbon-active-state.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Review New Note label', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
    document.body.replaceChildren();
  });

  it('reads Edit Note while the active cell has a note and New Note otherwise', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const { instance } = sheet;
    const ribbon = document.createElement('div');
    const button = document.createElement('button');
    button.dataset.ribbonCommand = 'newCommentReview';
    const label = document.createElement('span');
    label.textContent = en.ribbon.newComment;
    button.append(label);
    ribbon.append(button);
    document.body.append(ribbon);

    projectRibbonActiveState(ribbon, instance, { editNote: en.ribbon.editNote });
    expect([label.textContent, button.getAttribute('aria-label')]).toEqual([
      'New Note',
      'New Note',
    ]);

    setComment(instance.store, { sheet: 0, row: 0, col: 0 }, 'Check this');
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
    projectRibbonActiveState(ribbon, instance, { editNote: en.ribbon.editNote });
    expect([label.textContent, button.getAttribute('aria-label')]).toEqual([
      'Edit Note',
      'Edit Note',
    ]);

    mutators.setActive(instance.store, { sheet: 0, row: 1, col: 0 });
    projectRibbonActiveState(ribbon, instance, { editNote: en.ribbon.editNote });
    expect(label.textContent).toBe('New Note');
  });
});
