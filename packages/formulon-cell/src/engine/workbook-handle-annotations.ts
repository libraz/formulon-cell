import type { EngineCapabilities, PhoneticRun, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

type EngineCommentEntry = { row: number; col: number; author: string; text: string };

type CommentEnumerableWorkbook = Workbook & {
  getComments?: (
    sheet: number,
  ) => readonly EngineCommentEntry[] & { readonly status: { ok: boolean } };
};

/** Lossless hyperlink record returned by the native engine. Kept module-local
 * at the public package boundary; gallery commands use it without widening
 * the root cell API. */
export type EngineHyperlinkRecord = import('@libraz/formulon').HyperlinkEntry;

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleAnnotationsMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

export abstract class WorkbookHandleAnnotationsMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Read the cell's OOXML phonetic guide, if the current engine exposes it. */
  getCellPhonetic(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.phonetic) return null;
    const r = wb(this).getCellPhonetic(sheet, row, col);
    return r.status.ok && r.value ? r.value : null;
  }

  /** Set (or, with an empty string, clear) the cell's phonetic guide. The
   *  engine spans the whole cell text, so this replaces any per-run guide the
   *  cell carried — use `setCellPhoneticRuns` to preserve the spans. */
  setCellPhonetic(sheet: number, row: number, col: number, phonetic: string): boolean {
    assertAlive(this);
    if (!this.capabilities.phonetic) return false;
    return wb(this).setCellPhonetic(sheet, row, col, phonetic).ok;
  }

  /** Read the cell's phonetic guide span by span. Returns null when the engine
   *  has no per-run surface, which is distinct from the empty array an
   *  unannotated cell reports. */
  getCellPhoneticRuns(sheet: number, row: number, col: number): PhoneticRun[] | null {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return null;
    const r = wb(this).getCellPhoneticRuns(sheet, row, col);
    if (!r.status.ok) return null;
    return r.runs.map((run) => ({ start: run.sb, end: run.eb, text: run.text }));
  }

  /** Replace the cell's phonetic guide with `runs`, an ordered partition of the
   *  cell text. An empty array clears the guide. */
  setCellPhoneticRuns(
    sheet: number,
    row: number,
    col: number,
    runs: readonly PhoneticRun[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return false;
    return wb(this).setCellPhoneticRuns(
      sheet,
      row,
      col,
      runs.map((run) => ({ sb: run.start, eb: run.end, text: run.text })),
    ).ok;
  }

  /** Read the cell comment at `(sheet, row, col)`. Returns null when the
   *  cell has no comment or when the engine doesn't expose `getComment`. */
  getComment(sheet: number, row: number, col: number): { author: string; text: string } | null {
    assertAlive(this);
    if (!this.capabilities.comments) return null;
    const e = wb(this).getComment(sheet, row, col);
    return e ? { author: e.author, text: e.text } : null;
  }

  /** Snapshot every comment on `sheet` when the engine exposes a sheet-wide
   *  enumerator. Empty under stub or older engines. */
  getComments(sheet: number): { row: number; col: number; author: string; text: string }[] {
    assertAlive(this);
    if (!this.capabilities.commentsEnumerable) return [];
    const engineWb = wb(this) as CommentEnumerableWorkbook;
    if (typeof engineWb.getComments !== 'function') return [];
    const entries = engineWb.getComments(sheet);
    if (!entries.status.ok) return [];
    return entries.map((e) => ({
      row: e.row,
      col: e.col,
      author: e.author,
      text: e.text,
    }));
  }

  /** Persist a cell comment. Empty `text` removes it. No-op (returns false)
   *  under the stub. */
  setCommentEntry(sheet: number, row: number, col: number, author: string, text: string): boolean {
    assertAlive(this);
    if (!this.capabilities.comments) return false;
    const s = wb(this).setComment(sheet, row, col, author, text);
    return s.ok;
  }

  /** Full hyperlink snapshot, preserving rectangle and internal-link fields.
   * `null` means the native enumeration failed; an empty array is a valid
   * successful empty/capability-off result. */
  getHyperlinksFull(sheet: number): EngineHyperlinkRecord[] | null {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return [];
    const get = (
      wb(this) as unknown as {
        getHyperlinks?: (sheet: number) => {
          status: { ok: boolean };
          map: (
            mapper: (entry: EngineHyperlinkRecord) => EngineHyperlinkRecord,
          ) => EngineHyperlinkRecord[];
        };
      }
    ).getHyperlinks;
    if (typeof get !== 'function') return null;
    try {
      const result = get.call(wb(this), sheet);
      if (!result?.status?.ok || typeof result.map !== 'function') return null;
      return result.map((h) => ({
        row: h.row,
        col: h.col,
        lastRow: h.lastRow,
        lastCol: h.lastCol,
        target: h.target,
        location: h.location,
        display: h.display,
        tooltip: h.tooltip,
      }));
    } catch {
      return null;
    }
  }

  /** Snapshot of every hyperlink on `sheet`. Empty array under the stub or
   * native enumeration failure, preserving the historical narrow contract. */
  getHyperlinks(
    sheet: number,
  ): { row: number; col: number; target: string; display: string; tooltip: string }[] {
    const arr = this.getHyperlinksFull(sheet);
    if (!arr) return [];
    return arr.map((h) => ({
      row: h.row,
      col: h.col,
      target: h.target,
      display: h.display,
      tooltip: h.tooltip,
    }));
  }

  /** Whether the native lossless range hyperlink writer is available. */
  supportsHyperlinkRangeWrite(): boolean {
    assertAlive(this);
    return (
      this.capabilities.hyperlinks &&
      typeof (wb(this) as unknown as { addHyperlinkRange?: unknown }).addHyperlinkRange ===
        'function'
    );
  }

  /** Append a lossless hyperlink rectangle, including internal locations. */
  addHyperlinkRange(
    sheet: number,
    row: number,
    col: number,
    lastRow: number,
    lastCol: number,
    target: string,
    display = '',
    tooltip = '',
    location = '',
  ): boolean {
    assertAlive(this);
    if (!this.supportsHyperlinkRangeWrite()) return false;
    const add = (
      wb(this) as unknown as {
        addHyperlinkRange: (
          sheet: number,
          row: number,
          col: number,
          lastRow: number,
          lastCol: number,
          target: string,
          display: string,
          tooltip: string,
          location: string,
        ) => { ok: boolean };
      }
    ).addHyperlinkRange;
    try {
      const status = add.call(
        wb(this),
        sheet,
        row,
        col,
        lastRow,
        lastCol,
        target,
        display,
        tooltip,
        location,
      );
      return Boolean(status?.ok);
    } catch {
      return false;
    }
  }

  /** Append a hyperlink at `(sheet, row, col)`. Empty `display` / `tooltip`
   *  mean default. Returns false on engine failure or capability off. */
  addHyperlink(
    sheet: number,
    row: number,
    col: number,
    target: string,
    display = '',
    tooltip = '',
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    const s = wb(this).addHyperlink(sheet, row, col, target, display, tooltip, '');
    return s.ok;
  }

  /** Remove every hyperlink anchored at `(sheet, row, col)`. No-op when none
   *  match. Returns false on engine failure or capability off. */
  removeHyperlink(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    const s = wb(this).removeHyperlink(sheet, row, col);
    return s.ok;
  }

  /** Remove the hyperlink at `index` on `sheet`. */
  removeHyperlinkAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    const s = wb(this).removeHyperlinkAt(sheet, index);
    return s.ok;
  }

  /** Drop every hyperlink on `sheet`. */
  clearHyperlinks(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    const s = wb(this).clearHyperlinks(sheet);
    return s.ok;
  }
}

export function installAnnotationsMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleAnnotationsMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleAnnotationsMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
