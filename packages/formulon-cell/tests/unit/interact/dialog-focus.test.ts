import { afterEach, describe, expect, it } from 'vitest';
import { advanceDialogTab, focusableDialogItems } from '../../../src/interact/dialog-focus.js';

const radio = (name: string, checked = false): HTMLInputElement => {
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.checked = checked;
  return input;
};

describe('focusableDialogItems', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('keeps the active named radio at its own DOM position among other controls', () => {
    const root = document.createElement('div');
    const before = document.createElement('button');
    const first = radio('choice');
    const middle = document.createElement('button');
    const checked = radio('choice', true);
    const after = document.createElement('button');
    const outside = document.createElement('button');
    root.append(before, first, middle, checked, after);
    document.body.append(root, outside);
    outside.focus();

    expect(focusableDialogItems(root)).toEqual([before, middle, checked, after]);

    first.focus();
    expect(checked.checked).toBe(true);
    expect(focusableDialogItems(root)).toEqual([before, first, middle, after]);
  });

  it('uses checked or first eligible named radios and leaves unnamed radios independent', () => {
    const root = document.createElement('div');
    const unchecked = radio('selected');
    const unnamedFirst = radio('');
    const checked = radio('selected', true);
    const fallbackFirst = radio('fallback');
    const unnamedSecond = radio('');
    const fallbackSecond = radio('fallback');
    const outside = document.createElement('button');
    root.append(unchecked, unnamedFirst, checked, fallbackFirst, unnamedSecond, fallbackSecond);
    document.body.append(root, outside);
    outside.focus();

    expect(focusableDialogItems(root)).toEqual([
      unnamedFirst,
      checked,
      fallbackFirst,
      unnamedSecond,
    ]);
  });

  it('separates same-name groups by form and ignores hidden or disabled checked peers', () => {
    const root = document.createElement('div');
    const firstForm = document.createElement('form');
    const disabledChecked = radio('choice', true);
    disabledChecked.disabled = true;
    const firstFormEligible = radio('choice');
    firstForm.append(disabledChecked, firstFormEligible);

    const between = document.createElement('button');
    const secondForm = document.createElement('form');
    const hiddenWrapper = document.createElement('div');
    hiddenWrapper.hidden = true;
    const hiddenChecked = radio('choice', true);
    hiddenWrapper.appendChild(hiddenChecked);
    const secondFormEligible = radio('choice');
    secondForm.append(hiddenWrapper, secondFormEligible);

    const noFormFirst = radio('choice');
    const noFormChecked = radio('choice', true);
    const after = document.createElement('button');
    const outside = document.createElement('button');
    root.append(firstForm, between, secondForm, noFormFirst, noFormChecked, after);
    document.body.append(root, outside);
    outside.focus();

    expect(focusableDialogItems(root)).toEqual([
      firstFormEligible,
      between,
      secondFormEligible,
      noFormChecked,
      after,
    ]);
  });

  it('filters native hidden, disabled, and CSS-hidden controls but keeps opacity-zero controls', () => {
    const root = document.createElement('div');
    const visible = document.createElement('button');
    const opacityZero = document.createElement('button');
    opacityZero.style.opacity = '0';
    const hiddenInput = document.createElement('input');
    hiddenInput.type = 'hidden';
    hiddenInput.tabIndex = 0;
    const disabledFieldset = document.createElement('fieldset');
    disabledFieldset.disabled = true;
    const fieldsetInput = document.createElement('input');
    fieldsetInput.disabled = true;
    disabledFieldset.appendChild(fieldsetInput);
    const visibilityHidden = document.createElement('button');
    visibilityHidden.style.visibility = 'hidden';
    const visibilityCollapsed = document.createElement('button');
    visibilityCollapsed.style.visibility = 'collapse';
    const displayNone = document.createElement('button');
    displayNone.style.display = 'none';
    const hiddenAncestor = document.createElement('div');
    hiddenAncestor.style.display = 'none';
    const childOfHiddenAncestor = document.createElement('button');
    hiddenAncestor.appendChild(childOfHiddenAncestor);
    root.append(
      visible,
      opacityZero,
      hiddenInput,
      disabledFieldset,
      visibilityHidden,
      visibilityCollapsed,
      displayNone,
      hiddenAncestor,
    );
    document.body.appendChild(root);

    expect(fieldsetInput.matches(':disabled')).toBe(true);
    expect(focusableDialogItems(root)).toEqual([visible, opacityZero]);
  });

  it('does not consume modified or already-prevented Tab events', () => {
    const root = document.createElement('div');
    const first = document.createElement('button');
    const second = document.createElement('button');
    root.append(first, second);
    document.body.appendChild(root);
    first.focus();

    for (const modifier of ['altKey', 'ctrlKey', 'metaKey'] as const) {
      const event = new KeyboardEvent('keydown', {
        key: 'Tab',
        bubbles: true,
        cancelable: true,
        [modifier]: true,
      });
      expect(advanceDialogTab(root, event)).toBe(false);
      expect(event.defaultPrevented).toBe(false);
      expect(document.activeElement).toBe(first);
    }

    const prevented = new KeyboardEvent('keydown', {
      key: 'Tab',
      bubbles: true,
      cancelable: true,
    });
    prevented.preventDefault();
    expect(advanceDialogTab(root, prevented)).toBe(false);
    expect(prevented.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(first);
  });

  it('focuses the root and consumes Tab when no eligible controls remain', () => {
    const root = document.createElement('div');
    root.tabIndex = -1;
    const outside = document.createElement('button');
    const hidden = document.createElement('button');
    hidden.hidden = true;
    root.appendChild(hidden);
    document.body.append(root, outside);
    outside.focus();
    const event = new KeyboardEvent('keydown', {
      key: 'Tab',
      bubbles: true,
      cancelable: true,
    });

    expect(advanceDialogTab(root, event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(root);
  });
});
