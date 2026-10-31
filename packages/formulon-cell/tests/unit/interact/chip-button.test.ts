import { describe, expect, it } from 'vitest';
import {
  createInteractionButton,
  createInteractionChipButton,
} from '../../../src/interact/chip-button.js';

describe('interaction button factories', () => {
  it('creates the button directly in the requested owner document', () => {
    const ownerDocument = document.implementation.createHTMLDocument('embedded buttons');

    const button = createInteractionButton(
      {
        className: 'fc-test-button',
        text: 'Run',
        dataset: { fcAction: 'run' },
        role: 'menuitem',
      },
      ownerDocument,
    );

    expect(button.parentNode).toBeNull();
    expect(button.ownerDocument === ownerDocument).toBe(true);
    expect(button.textContent).toBe('Run');
    expect(button.dataset.fcAction).toBe('run');
    expect(button.getAttribute('role')).toBe('menuitem');
  });

  it('continues to use the global document when no owner is supplied', () => {
    const button = createInteractionButton({ className: 'fc-test-button', text: 'Run' });

    expect(button.ownerDocument).toBe(document);
  });

  it('forwards the owner document and chip options to the shared button factory', () => {
    const ownerDocument = document.implementation.createHTMLDocument('embedded chips');

    const chip = createInteractionChipButton(
      {
        className: 'fc-test-chip',
        label: 'Selection 1',
        dataset: { fcRange: 'first' },
        role: 'button',
        selected: true,
      },
      ownerDocument,
    );

    expect(chip.parentNode).toBeNull();
    expect(chip.ownerDocument === ownerDocument).toBe(true);
    expect(chip.className).toBe('fc-test-chip');
    expect(chip.textContent).toBe('Selection 1');
    expect(chip.dataset.fcRange).toBe('first');
    expect(chip.getAttribute('role')).toBe('button');
    expect(chip.getAttribute('aria-selected')).toBe('true');
  });
});
