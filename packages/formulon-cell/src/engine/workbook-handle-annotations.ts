import type {
  DrawingAnchorPatch,
  DrawingPlacementInput,
  DrawingProperties,
  Formulon,
  Hyperlink,
  HyperlinkInput,
  ImageInfo,
  Mention,
  MentionInput,
  Person,
  PersonInput,
  PhoneticAlignment,
  PhoneticType,
  PictureProperties,
  Sheet,
  ThreadedComment,
  ThreadedCommentInput,
} from '@libraz/formulon';
import { isFormulonError } from './result.js';
import type {
  Addr,
  EngineCapabilities,
  PhoneticProperties,
  PhoneticPropertiesRead,
  PhoneticRun,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  module: Formulon;
  wb: Workbook;
  assertAlive(): void;
};

/** A legacy note. `visible` is present only for a note shown without hovering. */
export interface EngineNoteEntry {
  /** The note's author; null when it has none. */
  author: string | null;
  text: string;
  visible?: true;
}

/** Lossless hyperlink record returned by the native engine. Kept module-local
 * at the public package boundary; gallery commands use it without widening
 * the root cell API. */
export type EngineHyperlinkRecord = Hyperlink;

/** The native threaded-comment record, kept as a named cell-layer export so
 * callers do not have to depend on the generated formulon declaration. */
export type EngineThreadedComment = ThreadedComment;
export type EngineThreadedCommentInput = ThreadedCommentInput;
export type EngineMention = Mention;
export type EngineMentionInput = MentionInput;
export type EnginePerson = Person;
export type EnginePersonInput = PersonInput;
export type EngineDrawingObject = DrawingProperties;
export type EngineImageInfo = ImageInfo;
export type EngineImage = PictureProperties;
export type EngineImageInsertOptions = DrawingPlacementInput;
export type EngineDrawingAnchorPatch = DrawingAnchorPatch;

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

function ws(handle: unknown, sheet: number): Sheet {
  return internals(handle).wb.sheet(sheet);
}

/** Run an engine write; false when the engine rejects it. */
function attempt(write: () => void): boolean {
  try {
    write();
    return true;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return false;
  }
}

/** Run an engine read; null when the engine rejects it. */
function attemptRead<T>(read: () => T): T | null {
  try {
    return read();
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return null;
  }
}

export abstract class WorkbookHandleAnnotationsMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Read the cell's OOXML phonetic guide, if the current engine exposes it. */
  getCellPhonetic(sheet: number, row: number, col: number): string | null {
    assertAlive(this);
    if (!this.capabilities.phonetic) return null;
    return attemptRead(() => ws(this, sheet).getCellPhonetic(row, col)) || null;
  }

  /** Set (or, with an empty string, clear) the cell's phonetic guide. The
   *  engine spans the whole cell text, so this replaces any per-run guide the
   *  cell carried — use `setCellPhoneticRuns` to preserve the spans. */
  setCellPhonetic(sheet: number, row: number, col: number, phonetic: string): boolean {
    assertAlive(this);
    if (!this.capabilities.phonetic) return false;
    return attempt(() =>
      ws(this, sheet).setCellPhonetic(row, col, phonetic.length > 0 ? phonetic : null),
    );
  }

  /** Read the cell's phonetic guide span by span. Returns null when the engine
   *  has no per-run surface, which is distinct from the empty array an
   *  unannotated cell reports. */
  getCellPhoneticRuns(sheet: number, row: number, col: number): PhoneticRun[] | null {
    assertAlive(this);
    if (!this.capabilities.phoneticRuns) return null;
    const runs = attemptRead(() => ws(this, sheet).getCellPhoneticRuns(row, col));
    if (!runs) return null;
    return runs.map((run) => ({ start: run.start, end: run.end, text: run.text }));
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
    return attempt(() =>
      ws(this, sheet).setCellPhoneticRuns(
        row,
        col,
        runs.map((run) => ({ start: run.start, end: run.end, text: run.text })),
      ),
    );
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
    const r = attemptRead(() => ws(this, sheet).getCellPhoneticProperties(row, col));
    if (!r) return null;
    return { fontId: r.fontIndex, type: r.type, alignment: r.alignment, present: r.stated };
  }

  /** Read the cell's "Show Phonetic" flag. Null when the engine has no such
   *  surface. */
  getCellShowPhonetic(sheet: number, row: number, col: number): boolean | null {
    assertAlive(this);
    if (!this.capabilities.showPhonetic) return null;
    return attemptRead(() => ws(this, sheet).getCellShowPhonetic(row, col));
  }

  /** Every stored cell on `sheet` whose "Show Phonetic" flag is on, blanks carrying only the flag
   *  included, row-major. `[]` when the engine has no such surface or refuses. */
  showPhoneticCells(sheet: number): readonly Addr[] {
    assertAlive(this);
    if (!this.capabilities.showPhonetic) return [];
    const cells = attemptRead(() => ws(this, sheet).getShowPhoneticCells());
    return cells ? cells.map((c) => ({ sheet, row: c.row, col: c.col })) : [];
  }

  /** Set the cell's "Show Phonetic" display flag. The engine keeps it on
   *  constants and blanks only, and it survives value writes. */
  setCellShowPhonetic(sheet: number, row: number, col: number, show: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.showPhonetic) return false;
    return attempt(() => ws(this, sheet).setCellShowPhonetic(row, col, show));
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
    return attempt(() =>
      ws(this, sheet).setCellPhoneticProperties(row, col, {
        fontIndex: properties.fontId,
        type: properties.type as PhoneticType,
        alignment: properties.alignment as PhoneticAlignment,
      }),
    );
  }

  /** Read the cell comment at `(sheet, row, col)`. Returns null when the
   *  cell has no comment or when the engine doesn't expose `getComment`. */
  getComment(sheet: number, row: number, col: number): EngineNoteEntry | null {
    assertAlive(this);
    if (!this.capabilities.comments) return null;
    const e = attemptRead(() => ws(this, sheet).getComment(row, col));
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
    return attemptRead(() => ws(this, sheet).getCommentVisible(row, col)) === true;
  }

  /** Show or hide an existing note; the engine keeps its box geometry. */
  setCommentVisible(sheet: number, row: number, col: number, visible: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.commentVisibility) return false;
    return attempt(() => ws(this, sheet).setCommentVisible(row, col, visible));
  }

  /** Snapshot every comment on `sheet`. Empty under the stub. */
  getComments(sheet: number): ({ row: number; col: number } & EngineNoteEntry)[] {
    assertAlive(this);
    if (!this.capabilities.commentsEnumerable) return [];
    const entries = attemptRead(() => ws(this, sheet).listComments());
    if (!entries) return [];
    return entries.map((e) => ({
      row: e.cell.row,
      col: e.cell.col,
      author: e.author,
      text: e.text,
      ...(this.getCommentVisible(sheet, e.cell.row, e.cell.col) ? { visible: true as const } : {}),
    }));
  }

  /** Persist a cell comment. Empty `text` removes it. `visible` restates the
   *  note's shown state, which a re-added note otherwise loses. No-op
   *  (returns false) under the stub. */
  setCommentEntry(
    sheet: number,
    row: number,
    col: number,
    author: string | null,
    text: string,
    visible?: boolean,
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.comments) return false;
    const written = attempt(() =>
      ws(this, sheet).setComment(row, col, text.length > 0 ? { text, author } : null),
    );
    if (!written) return false;
    if (visible === undefined || text.length === 0 || !this.capabilities.commentVisibility) {
      return true;
    }
    return attempt(() => ws(this, sheet).setCommentVisible(row, col, visible));
  }

  /** Read every native threaded comment on a sheet, including replies. Notes
   * are deliberately excluded: the two annotation models are independent in
   * OOXML and the engine rejects putting both on one cell. */
  getThreadedComments(sheet: number): EngineThreadedComment[] {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return [];
    const comments = attemptRead(() => ws(this, sheet).listThreadedComments());
    return comments
      ? comments.map((comment) => ({
          ...comment,
          mentions: comment.mentions.map((mention) => ({ ...mention })),
        }))
      : [];
  }

  /** Add a thread opener or reply. A read comment passes back as is; the
   * engine validates ids and mentions before mutating the workbook. */
  addThreadedComment(sheet: number, comment: EngineThreadedCommentInput): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return attempt(() => {
      ws(this, sheet).addThreadedComment({
        ...comment,
        mentions: comment.mentions?.map((mention) => ({ ...mention })),
      });
    });
  }

  editThreadedComment(
    sheet: number,
    id: string,
    text: string,
    mentions: readonly EngineMentionInput[] = [],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return attempt(() =>
      ws(this, sheet).updateThreadedComment(id, {
        text,
        mentions: mentions.map((mention) => ({ ...mention })),
      }),
    );
  }

  setThreadResolved(sheet: number, threadId: string, done: boolean): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return attempt(() => ws(this, sheet).updateThreadedComment(threadId, { done }));
  }

  removeThreadedComment(sheet: number, id: string): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return attempt(() => ws(this, sheet).removeThreadedComment(id));
  }

  getPersons(): EnginePerson[] {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return [];
    return wb(this)
      .listPeople()
      .map((person) => ({ ...person }));
  }

  addPerson(person: EnginePersonInput): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return attempt(() => {
      wb(this).addPerson({ ...person });
    });
  }

  removePerson(id: string): boolean {
    assertAlive(this);
    if (!this.capabilities.threadedComments) return false;
    return attempt(() => wb(this).removePerson(id));
  }

  /** Inspect image bytes before inserting them. This is useful to reject an
   * unsupported file before the UI creates a persistent object. */
  probeImage(bytes: Uint8Array): EngineImageInfo | null {
    assertAlive(this);
    if (!this.capabilities.images) return null;
    const info = attemptRead(() => internals(this).module.getImageInfo(bytes));
    return info ? { ...info } : null;
  }

  listDrawingObjects(sheet: number): EngineDrawingObject[] {
    assertAlive(this);
    if (!this.capabilities.drawings) return [];
    return (
      attemptRead(() =>
        ws(this, sheet)
          .listDrawings()
          .map((drawing) => drawing.getDrawingProperties()),
      ) ?? []
    );
  }

  getImage(sheet: number, objectId: number): EngineImage | null {
    assertAlive(this);
    if (!this.capabilities.images) return null;
    const image = attemptRead(() => ws(this, sheet).picture(objectId).getProperties());
    return image ? { ...image, bytes: Uint8Array.from(image.bytes) } : null;
  }

  insertImage(
    sheet: number,
    bytes: Uint8Array,
    options: EngineImageInsertOptions = {},
  ): number | null {
    assertAlive(this);
    if (!this.capabilities.images) return null;
    return attemptRead(() => ws(this, sheet).addPicture(bytes, { ...options }).objectId);
  }

  /** Remove the picture or chart `objectId` and the parts only it reached. */
  removeDrawing(sheet: number, objectId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.drawings) return false;
    return attempt(() => ws(this, sheet).removeDrawing(objectId));
  }

  /** Move and resize a picture or chart, keeping its object id. */
  setDrawingAnchor(sheet: number, objectId: number, placement: EngineDrawingAnchorPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.drawings) return false;
    return attempt(() =>
      ws(this, sheet)
        .drawing(objectId)
        .update({ anchor: { ...placement } }),
    );
  }

  /** Move the object to `index` of `listDrawingObjects` (0 is the back). */
  setDrawingZOrder(sheet: number, objectId: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.drawings) return false;
    return attempt(() => {
      const drawing = ws(this, sheet).drawing(objectId) as StackableDrawing;
      if (!drawing.setZOrder) throw new TypeError('drawing object has no stacking order');
      drawing.setZOrder(index);
    });
  }

  /** Capture a picture or chart as opaque bytes for `restoreDrawing`. */
  snapshotDrawing(sheet: number, objectId: number): Uint8Array | null {
    assertAlive(this);
    if (!this.capabilities.drawings) return null;
    const bytes = attemptRead(() => {
      const drawing = ws(this, sheet).drawing(objectId) as StackableDrawing;
      return drawing.snapshot ? drawing.snapshot() : null;
    });
    return bytes ? Uint8Array.from(bytes) : null;
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
    return attemptRead(
      () =>
        ws(this, sheet).restoreDrawing(bytes, {
          ...(options.newId !== undefined ? { assignNewId: options.newId } : {}),
        }).objectId,
    );
  }

  /** Full hyperlink snapshot, preserving the covered range and internal-link
   * fields. `null` means the native enumeration failed; an empty array is a
   * valid successful empty/capability-off result. */
  getHyperlinksFull(sheet: number): EngineHyperlinkRecord[] | null {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return [];
    const links = attemptRead(() => ws(this, sheet).listHyperlinks());
    if (!links) return null;
    return links.map((h) => ({
      range: { ...h.range },
      target: h.target,
      location: h.location,
      display: h.display,
      tooltip: h.tooltip,
    }));
  }

  /** Snapshot of every hyperlink on `sheet`. Empty array under the stub or
   * native enumeration failure, preserving the historical narrow contract. */
  getHyperlinks(
    sheet: number,
  ): { row: number; col: number; target: string; display: string; tooltip: string }[] {
    const arr = this.getHyperlinksFull(sheet);
    if (!arr) return [];
    return arr.map((h) => ({
      row: h.range.firstRow,
      col: h.range.firstCol,
      target: h.target ?? '',
      display: h.display ?? '',
      tooltip: h.tooltip ?? '',
    }));
  }

  /** Whether the native lossless range hyperlink writer is available. */
  supportsHyperlinkRangeWrite(): boolean {
    assertAlive(this);
    return this.capabilities.hyperlinks;
  }

  /** Set the lossless hyperlink rectangle anchored at `(row, col)`, including
   *  internal locations. Empty strings mean absent. */
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
    return attempt(() =>
      ws(this, sheet).setHyperlink(row, col, {
        ...hyperlinkInput(target, display, tooltip, location),
        range: { firstRow: row, firstCol: col, lastRow, lastCol },
      }),
    );
  }

  /** Set the hyperlink at `(sheet, row, col)`. Empty `display` / `tooltip`
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
    return attempt(() =>
      ws(this, sheet).setHyperlink(row, col, hyperlinkInput(target, display, tooltip, location)),
    );
  }

  /** Remove the hyperlink anchored at `(sheet, row, col)`. No-op when none
   *  matches. Returns false on engine failure or capability off. */
  removeHyperlink(sheet: number, row: number, col: number): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    return attempt(() => ws(this, sheet).setHyperlink(row, col, null));
  }

  /** Remove the hyperlink at `index` of `getHyperlinksFull(sheet)`, leaving others on the same
   *  anchor in place. */
  removeHyperlinkAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    return attempt(() => ws(this, sheet).removeHyperlinkAt(index));
  }

  /** Insert `record` at `index` of `getHyperlinksFull(sheet)`, replacing nothing; `index` equal to
   *  the count appends. Pairs with `removeHyperlinkAt` for a lossless undo. */
  insertHyperlinkAt(sheet: number, index: number, record: EngineHyperlinkRecord): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    return attempt(() =>
      ws(this, sheet).insertHyperlinkAt(index, {
        target: record.target,
        location: record.location,
        display: record.display,
        tooltip: record.tooltip,
        range: { ...record.range },
      }),
    );
  }

  /** Drop every hyperlink on `sheet`. */
  clearHyperlinks(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.hyperlinks) return false;
    return attempt(() => ws(this, sheet).clearHyperlinks());
  }
}

/** A drawing handle whose kind may support stacking and capture. */
type StackableDrawing = {
  setZOrder?: (zOrderIndex: number) => void;
  snapshot?: () => Uint8Array;
};

/** The engine's link record for formulon-cell's empty-means-absent strings. */
function hyperlinkInput(
  target: string,
  display: string,
  tooltip: string,
  location: string,
): HyperlinkInput {
  return {
    target: target.length > 0 ? target : null,
    location: location.length > 0 ? location : null,
    display: display.length > 0 ? display : null,
    tooltip: tooltip.length > 0 ? tooltip : null,
  };
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
