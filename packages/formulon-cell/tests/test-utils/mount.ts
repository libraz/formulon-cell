import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import type { MountOptions, SpreadsheetInstance } from '../../src/mount/types.js';
import { Spreadsheet } from '../../src/mount.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from './dom.js';

export interface MountedStubSheet {
  instance: SpreadsheetInstance;
  host: HTMLElement;
  workbook: WorkbookHandle;
  dispose: () => void;
}

/** Mount a `Spreadsheet` against a fresh host with the stub engine. Returns
 *  the instance, host, workbook and a single `dispose()` that:
 *  - calls `instance.dispose()` (idempotent in production code),
 *  - removes the host from `document.body`,
 *  - restores canvas / ResizeObserver stubs installed by this helper.
 *
 *  Callers should wrap in `beforeEach` / `afterEach`:
 *
 *  ```ts
 *  let sheet: MountedStubSheet;
 *  beforeEach(async () => { sheet = await mountStubSheet(); });
 *  afterEach(() => sheet.dispose());
 *  ```
 */
export async function mountStubSheet(
  opts: Omit<MountOptions, 'workbook'> & { workbook?: WorkbookHandle } = {},
): Promise<MountedStubSheet> {
  installDomStubs();
  const { host, cleanup } = createHostElement();
  const workbook = opts.workbook ?? (await WorkbookHandle.createDefault({ preferStub: true }));
  const instance = await Spreadsheet.mount(host, { ...opts, workbook });
  return {
    instance,
    host,
    workbook,
    dispose: () => {
      try {
        instance.dispose();
      } catch {
        // dispose is best-effort; tests assert lifecycle separately.
      }
      cleanup();
      uninstallDomStubs();
    },
  };
}
