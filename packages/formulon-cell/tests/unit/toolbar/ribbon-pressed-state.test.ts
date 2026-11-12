import { describe, expect, it } from 'vitest';
import { createRibbonButton, projectPressedState } from '../../../src/toolbar/ribbon/button.js';

describe('projectPressedState', () => {
  it('marks a ribbon toggle pressed and released in its class and aria-pressed', () => {
    const button = createRibbonButton({ className: 'fc-tb__rb' });
    projectPressedState(button, true);
    expect(button.classList.contains('fc-tb__rb--active')).toBe(true);
    expect(button.getAttribute('aria-pressed')).toBe('true');
    projectPressedState(button, false);
    expect(button.classList.contains('fc-tb__rb--active')).toBe(false);
    expect(button.getAttribute('aria-pressed')).toBe('false');
    expect(button.className).toBe('fc-tb__rb');
  });
});
