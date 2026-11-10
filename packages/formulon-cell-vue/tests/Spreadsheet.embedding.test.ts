import type { SpreadsheetInstance, Workbook } from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { type App, createApp, defineComponent, h, shallowReactive, shallowRef } from 'vue';
import { Spreadsheet, type SpreadsheetExposed } from '../src';
import { flush, stubWorkbook, waitFor } from './test-utils/mount';

const apps: App[] = [];
const workbooks: Workbook[] = [];

afterEach(() => {
  for (const app of apps.splice(0)) app.unmount();
  for (const workbook of workbooks.splice(0)) workbook.dispose();
  document.body.replaceChildren();
});

const freshWorkbook = async (): Promise<Workbook> => {
  const workbook = await stubWorkbook();
  workbooks.push(workbook);
  return workbook;
};

const mountEmbedded = async (initial: Record<string, unknown>) => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const props = shallowReactive<Record<string, unknown>>({ locale: 'en', ...initial });
  const onReady = vi.fn();
  const onError = vi.fn();
  const onUpdate = vi.fn();
  const child = shallowRef<SpreadsheetExposed | null>(null);
  const app = createApp(
    defineComponent({
      render: () =>
        h(Spreadsheet, {
          ...props,
          onReady,
          onError,
          onUpdate,
          ref: (el: unknown) => {
            child.value = el as SpreadsheetExposed | null;
          },
        }),
    }),
  );
  app.mount(host);
  apps.push(app);
  await waitFor(() => onReady.mock.calls.length > 0);
  const instance = child.value?.spreadsheet as SpreadsheetInstance;
  return { props, onReady, onError, onUpdate, instance, host };
};

describe('<Spreadsheet> embedding props', () => {
  it('passes mount-time options through, including seed', async () => {
    const seed = vi.fn((workbook: Workbook) => {
      workbook.sheet(0).setNumber(0, 0, 5);
    });
    const { instance } = await mountEmbedded({ seed, theme: 'ink', toolbar: ribbon() });
    expect(seed).toHaveBeenCalledTimes(1);
    expect(instance.workbook.sheet(0).getTypedValue(0, 0)).toMatchObject({
      number: 5,
    });
  });

  it('shows the ribbon for a toolbar prop and removes it when the prop is cleared', async () => {
    const { props, host, onUpdate } = await mountEmbedded({
      workbook: await freshWorkbook(),
      toolbar: ribbon(),
    });
    expect(host.querySelector('[data-ribbon-tab]')).not.toBeNull();
    props.toolbar = null;
    await flush();
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['toolbar'] });
    expect(host.querySelector('[data-ribbon-tab]')).toBeNull();
  });

  it('resets a key to its default when the prop is removed', async () => {
    const { props, onUpdate } = await mountEmbedded({
      workbook: await freshWorkbook(),
      policy: { allow: false } as never,
    });
    onUpdate.mockClear();
    props.policy = undefined;
    await flush();
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['policy'] });
  });

  it('reports a rejected workbook replacement and keeps the instance on its workbook', async () => {
    const original = await freshWorkbook();
    const { props, instance, onError } = await mountEmbedded({ workbook: original });
    const doomed = await stubWorkbook();
    doomed.dispose();
    props.workbook = doomed;
    await flush();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(instance.disposed).toBe(false);
    expect(instance.workbook).toBe(original);
  });

  it('keeps applying the remaining keys after one is rejected', async () => {
    const { props, instance, onError, onUpdate } = await mountEmbedded({
      workbook: await freshWorkbook(),
    });
    const doomed = await stubWorkbook();
    doomed.dispose();
    onUpdate.mockClear();
    props.workbook = doomed;
    props.locale = 'ja';
    await flush();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(instance.i18n.locale).toBe('ja');
    expect(onUpdate).toHaveBeenCalledWith({ keys: ['locale'] });
  });

  it('keeps the strings overlay when the locale changes through the prop', async () => {
    const { props, instance } = await mountEmbedded({
      workbook: await freshWorkbook(),
      strings: { ribbon: { tabs: { home: 'Start' } } } as never,
    });
    props.locale = 'ja';
    await flush();
    expect(instance.i18n.locale).toBe('ja');
    expect(instance.disposed).toBe(false);
  });
});
