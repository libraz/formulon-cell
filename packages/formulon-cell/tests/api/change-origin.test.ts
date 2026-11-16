/** The workbook `change` event reports the origin of the write that caused it. */
import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en' });
  live.push(instance);
  return instance;
};

afterEach(() => {
  while (live.length) live.pop()?.dispose();
});

describe('workbook change origin', () => {
  it('reports the origin passed to applyChanges once the task settles', async () => {
    const sheet = await mountSheet();
    const changes = vi.fn();
    sheet.workbook.on('change', changes);
    sheet.applyChanges([{ at: 'A1', input: '5' }], { history: 'reset', origin: 'host:import' });
    await Promise.resolve();
    expect(changes).toHaveBeenCalledTimes(1);
    expect(changes.mock.calls[0]?.[0]).toMatchObject({ origin: 'host:import' });
  });

  it('keeps each origin when two applyChanges calls share a task', async () => {
    const sheet = await mountSheet();
    const origins: string[] = [];
    sheet.workbook.on('change', (e) => origins.push(e.origin));
    sheet.applyChanges([{ at: 'A1', input: '1' }], { history: 'record', origin: 'host:a' });
    sheet.applyChanges([{ at: 'A2', input: '2' }], { history: 'record', origin: 'host:b' });
    await Promise.resolve();
    expect(origins).toEqual(['host:a', 'host:b']);
  });
});
