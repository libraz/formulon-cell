import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import en from '../../../src/i18n/en.js';
import type { Strings } from '../../../src/i18n/strings.js';
import { prepareMountHost, releaseMountHost } from '../../../src/mount/host.js';

describe('mount/host', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    host.remove();
  });

  describe('prepareMountHost', () => {
    it('applies the spreadsheet a11y attributes and theme class', () => {
      prepareMountHost(host, en as Strings, 'paper');

      expect(host.classList.contains('fc-host')).toBe(true);
      expect(host.getAttribute('tabindex')).toBe('0');
      expect(host.getAttribute('role')).toBe('region');
      expect(host.getAttribute('aria-roledescription')).toBe('spreadsheet');
      expect(host.getAttribute('aria-label')).toBe((en as Strings).a11y.spreadsheet);
      expect(host.dataset.fcTheme).toBe('paper');
    });

    it('defaults theme to paper when undefined', () => {
      prepareMountHost(host, en as Strings, undefined);
      expect(host.dataset.fcTheme).toBe('paper');
    });

    it('honours the requested theme', () => {
      prepareMountHost(host, en as Strings, 'ink');
      expect(host.dataset.fcTheme).toBe('ink');
    });

    it('clears any pre-existing children of the host', () => {
      const stale = document.createElement('span');
      stale.textContent = 'old';
      host.appendChild(stale);

      prepareMountHost(host, en as Strings, 'paper');
      expect(host.children.length).toBe(0);
    });

    it('assigns a unique instance id that increments per mount', () => {
      const first = prepareMountHost(host, en as Strings, 'paper');
      const other = document.createElement('div');
      document.body.appendChild(other);
      const second = prepareMountHost(other, en as Strings, 'paper');

      expect(first).toMatch(/^fc-\d+$/);
      expect(second).toMatch(/^fc-\d+$/);
      expect(first).not.toBe(second);
      other.remove();
    });
  });

  describe('releaseMountHost', () => {
    it('clears host content, removes the fc-host class, and drops the instance id', () => {
      const id = prepareMountHost(host, en as Strings, 'paper');
      host.dataset.fcEngineState = 'ready';
      host.appendChild(document.createElement('div'));

      releaseMountHost(host, id);

      expect(host.children.length).toBe(0);
      expect(host.classList.contains('fc-host')).toBe(false);
      expect(host.dataset.fcInstId).toBeUndefined();
      expect(host.dataset.fcEngineState).toBeUndefined();
    });

    it('is a no-op when the instance id no longer matches', () => {
      const id = prepareMountHost(host, en as Strings, 'paper');
      host.appendChild(document.createElement('div'));
      host.dataset.fcInstId = 'fc-someone-else';

      releaseMountHost(host, id);

      expect(host.children.length).toBe(1);
      expect(host.classList.contains('fc-host')).toBe(true);
      expect(host.dataset.fcInstId).toBe('fc-someone-else');
    });
  });
});
