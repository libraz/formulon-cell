import { Window as HappyWindow } from 'happy-dom';
import { afterEach, describe, expect, it } from 'vitest';
import { isSubmitEnter } from '../../../src/interact/mac-dialog-keys.js';

describe('isSubmitEnter', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('keeps native Enter on a foreign-document button and accepts it on an input', async () => {
    const ownerWindow = new HappyWindow();
    const ownerDocument = ownerWindow.document as unknown as Document;
    const ownerView = ownerDocument.defaultView;
    if (!ownerView) {
      await ownerWindow.happyDOM.abort();
      throw new Error('foreign document has no defaultView');
    }
    const button = ownerDocument.createElement('button');
    const input = ownerDocument.createElement('input');
    ownerDocument.body.append(button, input);
    const submitResults: boolean[] = [];
    button.addEventListener('keydown', (event) => {
      submitResults.push(isSubmitEnter(event as KeyboardEvent));
    });
    input.addEventListener('keydown', (event) => {
      submitResults.push(isSubmitEnter(event as KeyboardEvent));
    });
    const keyEvent = (): KeyboardEvent =>
      new ownerView.KeyboardEvent('keydown', {
        key: 'Enter',
        bubbles: true,
        cancelable: true,
      });
    const realmTargetEvent = (element: Element): KeyboardEvent =>
      ({
        key: 'Enter',
        isComposing: false,
        target: {
          nodeType: element.nodeType,
          closest: element.closest.bind(element),
        },
      }) as unknown as KeyboardEvent;

    try {
      expect(button.ownerDocument).not.toBe(document);
      button.dispatchEvent(keyEvent());
      input.dispatchEvent(keyEvent());

      expect(submitResults).toEqual([false, true]);
      expect(isSubmitEnter(realmTargetEvent(button))).toBe(false);
      expect(isSubmitEnter(realmTargetEvent(input))).toBe(true);
    } finally {
      await ownerWindow.happyDOM.abort();
    }
  });

  it('ignores the IME confirmation Enter that arrives with keyCode 229 after composition', () => {
    const input = document.createElement('input');
    document.body.append(input);
    const results: boolean[] = [];
    input.addEventListener('keydown', (event) => results.push(isSubmitEnter(event)));

    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', keyCode: 229, bubbles: true }),
    );
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));

    expect(results).toEqual([false, true]);
  });
});
