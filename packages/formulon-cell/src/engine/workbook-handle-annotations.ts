import type {
  DrawingObject,
  ImageInfoResult,
  ImageInsertOptions,
  ImagePlacement,
  ImageResult,
  Mention,
  Person,
  ThreadedComment,
} from '@libraz/formulon';
import type {
  EngineCapabilities,
  PhoneticProperties,
  PhoneticPropertiesRead,
  PhoneticRun,
  Workbook,
} from './types.js';
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

/** A legacy note. `visible` is present only for a note shown without hovering. */
export interface EngineNoteEntry {
  author: string;
  text: string;
  visible?: true;
}

/** Lossless hyperlink record returned by the native engine. Kept module-local
 * at the public package boundary; gallery commands use it without widening
 * the root cell API. */
export type EngineHyperlinkRecord = import('@libraz/formulon').HyperlinkEntry;

/** The native threaded-comment record, kept as a named cell-layer export so
 * callers do not have to depend on the generated formulon declaration. */
export type EngineThreadedComment = ThreadedComment;
export type EngineMention = Mention;
export type EnginePerson = Person;
export type EngineDrawingObject = DrawingObject;
export type EngineImageInfo = ImageInfoResult;
export type EngineImage = ImageResult;
export type EngineImageInsertOptions = ImageInsertOptions;

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

  /** Read how the cell's guide renders. Returns null when the engine has no
   *  property surface. */
  getCellPhoneticProperties(
    sheet: number,
    row: number,
    col: number,
  ): PhoneticPropertiesRead | null {
    assertAlive(this);
    if (!this.capabilities.phoneticProperties) return null;
    const r = wb(this).getCellPhoneticProperties(sheet, row, col);
    if (!r.status.ok) return null;
    return { fontId: r.fontId, type: r.type, alignment: r.alignment, present: r.present };
  }

  /** Read the cell's "Show Phonetic" flag. Null when the engine has no such
   *  surface. */
  getCellShowPhonetic(sheet: number, row: number, col: number): boolean | null {
    assertAlive(this);
    if (!this.capabilities.showPhonetic) return null;
    const r = wb(this).getCellShowPhonetic(sheet, row, col);
    return r.status.ok ? r.show : null;
  }

  /** Set the cell's "Show Phonetic" display flag. The engine keeps it on
   *  constants and blanks only, and it survives value writes. */
  setCellShowPhonetic(sheet: number, row: number, col: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.showPhonetic) return false;
    return wb(this).setCellShowPhonetic(sheet, row, col, show).ok;
  }

  /** Store how the cell's guide renders. The engine drops the properties when
   *  the cell's value is written again, so call this after the value. */
  setCellPhoneticProperties(
    sheet: number,
    row: number,
    col: number,
    properties: PhoneticProperties,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.phoneticProperties) return false;
    return wb(this).setCellPhoneticProperties(sheet, row, col, {
      fontId: properties.fontId,
      type: properties.type,
      alignment: properties.alignment,
    }).ok;
  }

  /** Read the cell comment at `(sheet, row, col)`. Returns null when the
   *  cell has no comment or when the engine doesn't expose `getComment`. */
  getComment(sheet: number, row: number, col: number): EngineNoteEntry | null {
    assertAlive(this);
    if (!this.capabilities.comments) return null;
    const e = wb(this).getComment(sheet, row, col);
    if (!e) return null;
    return this.getCommentVisible(sheet, row, col)
      ? { author: e.author, text: e.text, visible: true }
      : { author: e.author, text: e.text };
  }

  /** Whether the note is shown without hovering. False when absent or when
   *  the engine cannot report visibility. */
  getCommentVisible(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.commentVisibility) return false;
    const r = wb(this).getCommentVisible(sheet, row, col);
    return r.status.ok && r.value === 1;
  }

  /** Show or hide an existing note; the engine keeps its box geometry. */
  setCommentVisible(sheet: number, row: number, col: number, visible: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.commentVisibility) return false;
    return wb(this).setCommentVisible(sheet, row, col, visible).ok;
  }

  /** Snapshot every comment on `sheet` when the engine exposes a sheet-wide
   *  enumerator. Empty under stub or older engines. */
  getComments(sheet: number): ({ row: number; col: number } & EngineNoteEntry)[] {
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
      ...(this.getCommentVisible(sheet, e.row, e.col) ? { visible: true as const } : {}),
    }));
  }

  /** Persist a cell comment. Empty `text` removes it. `visible` restates the
   *  note's shown state, which a re-added note otherwise loses. No-op
   *  (returns false) under the stub. */
  setCommentEntry(
    sheet: number,
    row: number,
    col: number,
    author: string,
    text: string,
    visible?: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.comments) return false;
    const s = wb(this).setComment(sheet, row, col, author, text);
    if (!s.ok) return false;
    if (visible === undefined || text.length === 0 || !this.capabilities.commentVisibility) {
      return true;
    }
    return wb(this).setCommentVisible(sheet, row, col, visible).ok;
  }

  /** Read every native threaded comment on a sheet, including replies. Notes
   * are deliberately excluded: the two annotation models are independent in
   * OOXML and the engine rejects putting both on one cell. */
  getThreadedComments(sheet: number): EngineThreadedComment[] {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return [];
    const result = wb(this).getThreadedComments(sheet);
    return result.status.ok
      ? result.comments.map((comment) => ({
          ...comment,
          mentions: comment.mentions.map((mention) => ({ ...mention })),
        }))
      : [];
  }

  /** Add a thread opener or reply. The caller supplies a stable GUID and may
   * include mentions; the engine validates both before mutating the workbook. */
  addThreadedComment(
    sheet: number,
    comment: Partial<EngineThreadedComment> & Pick<EngineThreadedComment, 'id' | 'text'>,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return wb(this).addThreadedComment(sheet, {
      ...comment,
      mentions: comment.mentions?.map((mention) => ({ ...mention })),
    }).ok;
  }

  editThreadedComment(
    sheet: number,
    id: string,
    text: string,
    mentions: readonly EngineMention[] = [],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return wb(this).editThreadedComment(
      sheet,
      id,
      text,
      mentions.map((mention) => ({ ...mention })),
    ).ok;
  }

  setThreadResolved(sheet: number, threadId: string, done: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return wb(this).setThreadResolved(sheet, threadId, done).ok;
  }

  removeThreadedComment(sheet: number, id: string): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return wb(this).removeThreadedComment(sheet, id).ok;
  }

  getPersons(): EnginePerson[] {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return [];
    const result = wb(this).getPersons();
    return result.status.ok ? result.persons.map((person) => ({ ...person })) : [];
  }

  addPerson(person: EnginePerson): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return wb(this).addPerson({ ...person }).ok;
  }

  removePerson(id: string): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return wb(this).removePerson(id).ok;
  }

  /** Inspect image bytes before inserting them. This is useful to reject an
   * unsupported file before the UI creates a persistent object. */
  probeImage(bytes: Uint8Array): EngineImageInfo | null {
    assertAlive(this);
    if (!this.capabilities.images) return null;
    const result = wb(this).probeImage(bytes);
    return result.status.ok ? { ...result } : null;
  }

  listDrawingObjects(sheet: number): EngineDrawingObject[] {
    assertAlive(this);
    if (!this.capabilities.drawings) return [];
    const result = wb(this).listDrawingObjects(sheet);
    return result.status.ok ? result.map((object) => ({ ...object })) : [];
  }

  getImage(sheet: number, objectId: number): EngineImage | null {
    assertAlive(this);
    if (!this.capabilities.images) return null;
    const result = wb(this).getImage(sheet, objectId);
    return result.status.ok ? { ...result, bytes: Uint8Array.from(result.bytes) } : null;
  }

  insertImage(
    sheet: number,
    bytes: Uint8Array,
    options: EngineImageInsertOptions = {},
  ): number | null {
    assertAlive(this);
    if (!this.capabilities.images) return null;
    const result = wb(this).insertImage(sheet, bytes, { ...options });
    return result.status.ok ? result.objectId : null;
  }

  /** Remove the picture or chart `objectId` and the parts only it reached. */
  removeDrawing(sheet: number, objectId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.drawings) return false;
    return wb(this).removeDrawing(sheet, objectId).ok;
  }

  /** Move and resize a picture or chart, keeping its object id. */
  setDrawingAnchor(sheet: number, objectId: number, placement: ImagePlacement): boolean {
    assertAlive(this);
    if (!this.capabilities.drawings) return false;
    return wb(this).setDrawingAnchor(sheet, objectId, { ...placement }).ok;
  }

  /** Move the object to `index` of `listDrawingObjects` (0 is the back). */
  setDrawingZOrder(sheet: number, objectId: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.drawings) return false;
    return wb(this).setDrawingZOrder(sheet, objectId, index).ok;
  }

  /** Capture a picture or chart as opaque bytes for `restoreDrawing`. */
  snapshotDrawing(sheet: number, objectId: number): Uint8Array | null {
    assertAlive(this);
    if (!this.capabilities.drawings) return null;
    const result = wb(this).snapshotDrawing(sheet, objectId);
    return result.status.ok ? Uint8Array.from(result.bytes) : null;
  }

  /** Put a `snapshotDrawing` capture back and return the object id: the same
   *  id in place by default, or a fresh one on top with `newId`. */
  restoreDrawing(
    sheet: number,
    bytes: Uint8Array,
    options: { newId?: boolean } = {},
  ): number | null {
    assertAlive(this);
    if (!this.capabilities.drawings) return null;
    const result = wb(this).restoreDrawing(sheet, bytes, { ...options });
    return result.status.ok ? result.objectId : null;
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
    location = '',
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    const s = wb(this).addHyperlink(sheet, row, col, target, display, tooltip, location);
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
