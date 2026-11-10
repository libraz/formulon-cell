import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { useI18n, useSelection, useSpreadsheet } from '../src/hooks';
import { type MountedReactSpreadsheet, mountReactSpreadsheet } from './test-utils/mount';

const Probe = ({ inst, n }: { inst: SpreadsheetInstance; n: number }) => {
  useSelection(inst);
  useI18n(inst);
  useSpreadsheet(inst, (s) => s.selection.get().active.row + n * 0);
  return null;
};

describe('hooks subscription', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('subscribes once per instance across re-renders and unsubscribes on unmount', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance;
    const onInstance = vi.spyOn(inst, 'on');
    const onSelection = vi.spyOn(inst.selection, 'on');
    const onI18n = vi.spyOn(inst.i18n, 'on');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    for (let n = 0; n < 3; n += 1) {
      await act(async () => {
        root.render(<Probe inst={inst} n={n} />);
      });
    }
    // useSpreadsheet listens to five instance events, useSpreadsheet and useSelection to the
    // selection, useSpreadsheet and useI18n to the locale; none grows with the render count.
    expect(onInstance).toHaveBeenCalledTimes(5);
    expect(onSelection).toHaveBeenCalledTimes(2);
    expect(onI18n).toHaveBeenCalledTimes(2);

    await act(async () => root.unmount());
    host.remove();
  });
});
