import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

import { useSelection, useSpreadsheet } from '../src/hooks';
import {
  installReactDomStubs,
  type MountedReactSpreadsheet,
  mountReactSpreadsheet,
  uninstallReactDomStubs,
} from './test-utils/mount';

const Probe = ({ inst, n }: { inst: SpreadsheetInstance; n: number }) => {
  useSelection(inst);
  useSpreadsheet(inst, (s) => s.selection.active.row + n * 0, 0);
  return null;
};

describe('hooks store subscription', () => {
  let mounted: MountedReactSpreadsheet | null = null;

  afterEach(async () => {
    if (mounted) await mounted.dispose();
    mounted = null;
    document.body.replaceChildren();
  });

  it('subscribes once per instance across re-renders', async () => {
    mounted = await mountReactSpreadsheet();
    const inst = mounted.instance as SpreadsheetInstance;
    const subscribe = vi.spyOn(inst.store, 'subscribe');
    installReactDomStubs();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    for (let n = 0; n < 3; n += 1) {
      await act(async () => {
        root.render(<Probe inst={inst} n={n} />);
      });
    }
    // One subscription per hook, regardless of render count.
    expect(subscribe).toHaveBeenCalledTimes(2);
    await act(async () => root.unmount());
    host.remove();
    uninstallReactDomStubs();
  });
});
