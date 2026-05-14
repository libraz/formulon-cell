import { afterEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../src/mount.js';

describe('Spreadsheet.mount error handling', () => {
  afterEach(() => {
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it('renders the built-in mount error and rejects when the engine cannot start', async () => {
    const err = new Error('SharedArrayBuffer is missing');
    vi.spyOn(WorkbookHandle, 'createDefault').mockRejectedValue(err);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const onError = vi.fn();

    await expect(Spreadsheet.mount(host, { onError })).rejects.toThrow(err);

    expect(onError).toHaveBeenCalledWith(err);
    expect(host.dataset.fcEngineState).toBe('error');
    expect(host.querySelector('.fc-mount-error')).toBeTruthy();
    expect(host.textContent).toContain('Spreadsheet engine unavailable');
    expect(host.textContent).toContain('SharedArrayBuffer is missing');
  });

  it('can suppress the built-in panel for framework-native fallbacks', async () => {
    const err = new Error('boom');
    vi.spyOn(WorkbookHandle, 'createDefault').mockRejectedValue(err);
    const host = document.createElement('div');
    document.body.appendChild(host);

    await expect(Spreadsheet.mount(host, { renderError: false })).rejects.toThrow(err);

    expect(host.dataset.fcEngineState).toBe('error');
    expect(host.querySelector('.fc-mount-error')).toBeNull();
  });
});
