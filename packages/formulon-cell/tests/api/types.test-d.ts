import type { Value } from '@libraz/formulon';
import type {
  BatchRejection,
  ChangeBatchResult,
  CommandId,
  CommandIntent,
  DialogId,
  Disposer,
  InteractionOperation,
  InteractionPolicy,
  SaveResult,
  Sheet,
  SheetPrintDocument,
  SpreadsheetFailure,
  SpreadsheetInstance,
  SpreadsheetOptions,
  UpdateEvent,
  Workbook,
} from '@libraz/formulon-cell';
import type { saveWithPrompts } from '@libraz/formulon-cell/toolbar';
import { describe, expectTypeOf, it } from 'vitest';

/** True when `T` admits `undefined`. */
type AdmitsUndefined<T> = undefined extends T ? true : false;

declare const wb: Workbook;
declare const sheet: Sheet;
declare const instance: SpreadsheetInstance;

describe('types', () => {
  it('every on() returns a Disposer', () => {
    expectTypeOf<Disposer>().toEqualTypeOf<() => void>();
    expectTypeOf(wb.on('change', () => {})).toEqualTypeOf<Disposer>();
    expectTypeOf(instance.on('update', () => {})).toEqualTypeOf<Disposer>();
    expectTypeOf(instance.selection.on('change', () => {})).toEqualTypeOf<Disposer>();
    expectTypeOf(instance.history.on('change', () => {})).toEqualTypeOf<Disposer>();
    expectTypeOf(instance.i18n.on('change', () => {})).toEqualTypeOf<Disposer>();
  });

  it('commands.can takes a command id or an intent', () => {
    expectTypeOf(instance.commands.can('bold')).toEqualTypeOf<boolean>();
    expectTypeOf(instance.commands.can({ kind: 'copy' })).toEqualTypeOf<boolean>();
    expectTypeOf(
      instance.commands.can({ kind: 'edit', operation: 'format', sheet: 'Sheet1', at: 'A1:B2' }),
    ).toEqualTypeOf<boolean>();
    expectTypeOf<CommandIntent>().toMatchTypeOf<Parameters<typeof instance.commands.can>[0]>();
  });

  it('options read back as a readonly snapshot and updates name their source', () => {
    expectTypeOf(instance.options).toEqualTypeOf<Readonly<SpreadsheetOptions>>();
    expectTypeOf(instance.options.showAllNotes).toEqualTypeOf<boolean | undefined>();
    expectTypeOf<UpdateEvent['source']>().toEqualTypeOf<'api' | 'user'>();
  });

  it('absence is null, never undefined', () => {
    expectTypeOf<AdmitsUndefined<SpreadsheetInstance['clipboard']>>().toEqualTypeOf<false>();
    expectTypeOf<AdmitsUndefined<SpreadsheetInstance['formatPainter']>>().toEqualTypeOf<false>();
    expectTypeOf<AdmitsUndefined<SpreadsheetInstance['borderDraw']>>().toEqualTypeOf<false>();
    expectTypeOf<AdmitsUndefined<SpreadsheetInstance['toolbar']>>().toEqualTypeOf<false>();
    expectTypeOf<
      AdmitsUndefined<ReturnType<SpreadsheetInstance['cells']['get']>>
    >().toEqualTypeOf<false>();
    expectTypeOf<
      AdmitsUndefined<ReturnType<SpreadsheetInstance['formulaHints']['get']>>
    >().toEqualTypeOf<false>();
  });

  it('values are the engine Value', () => {
    expectTypeOf(sheet.getTypedValue('A1')).toEqualTypeOf<Value>();
    expectTypeOf(sheet.evaluateTypedFormula('A1', '=1')).toEqualTypeOf<Value>();
  });

  it('save returns a SaveResult with an ArrayBuffer-backed byte array', () => {
    expectTypeOf(wb.save()).toEqualTypeOf<SaveResult>();
    expectTypeOf<SaveResult['bytes']>().toEqualTypeOf<Uint8Array<ArrayBuffer>>();
    expectTypeOf<SaveResult['format']>().toEqualTypeOf<'xlsx' | 'xlsb'>();
  });

  it('the failure union is closed', () => {
    expectTypeOf<SpreadsheetFailure>().toEqualTypeOf<
      | 'disposed'
      | 'feature-disabled'
      | 'policy-denied'
      | 'protected'
      | 'limit-exceeded'
      | 'workbook-in-use'
    >();
  });

  it('command and dialog ids reject unknown names', () => {
    expectTypeOf<'insertRows'>().toMatchTypeOf<CommandId>();
    expectTypeOf<'format'>().toMatchTypeOf<DialogId>();
    // @ts-expect-error not a command id
    instance.commands.execute('insertRowsAbove');
    // @ts-expect-error not a dialog id
    instance.dialogs.open('noSuchDialog');
  });

  it('applyChanges requires a history mode', () => {
    // @ts-expect-error history is required
    instance.applyChanges([]);
    // @ts-expect-error history is required
    instance.applyChanges([], {});
    instance.applyChanges([], { history: 'record' });
  });

  it('protect and unprotect are async', () => {
    expectTypeOf(sheet.protect()).toEqualTypeOf<Promise<void>>();
    expectTypeOf(sheet.unprotect()).toEqualTypeOf<Promise<void>>();
  });

  it('history repeats like it undoes and redoes', () => {
    expectTypeOf(instance.history.canRepeat).toEqualTypeOf<boolean>();
    expectTypeOf(instance.history.canUndo).toEqualTypeOf<boolean>();
    expectTypeOf(instance.history.repeat()).toEqualTypeOf<boolean>();
  });

  it('printDocument returns the print document', () => {
    expectTypeOf(instance.printDocument()).toEqualTypeOf<SheetPrintDocument>();
    expectTypeOf<SheetPrintDocument['html']>().toEqualTypeOf<string>();
  });

  it('saveWithPrompts takes the spreadsheet and optional format and host', () => {
    expectTypeOf<Parameters<typeof saveWithPrompts>[0]>().toEqualTypeOf<SpreadsheetInstance>();
    expectTypeOf<NonNullable<Parameters<typeof saveWithPrompts>[1]>>().toEqualTypeOf<{
      format?: 'xlsx' | 'xlsb';
      host?: HTMLElement;
    }>();
  });

  it('keeps batches all or nothing: no partial result, skip mode or moveCells operation', () => {
    type FacadeStatus = ReturnType<SpreadsheetInstance['applyChanges']>['status'];
    expectTypeOf<FacadeStatus>().toEqualTypeOf<'applied' | 'refused'>();
    expectTypeOf<ReturnType<Sheet['importCsv']>['status']>().toEqualTypeOf<FacadeStatus>();
    expectTypeOf<ChangeBatchResult['status']>().toEqualTypeOf<'applied' | 'refused'>();
    expectTypeOf<Extract<InteractionOperation, 'moveCells'>>().toEqualTypeOf<never>();
    // One shape for every batch of cell writes, refusals included.
    expectTypeOf<
      ReturnType<SpreadsheetInstance['applyChanges']>
    >().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<Sheet['importCsv']>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<SpreadsheetInstance['importCsv']>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ReturnType<Sheet['replaceAll']>>().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<ChangeBatchResult['rejections'][number]>().toEqualTypeOf<BatchRejection>();
    expectTypeOf<keyof ChangeBatchResult>().toEqualTypeOf<
      'status' | 'applied' | 'rejections' | 'range'
    >();
    expectTypeOf<
      ReturnType<NonNullable<SpreadsheetInstance['clipboard']>['pasteSpecial']>
    >().toEqualTypeOf<ChangeBatchResult>();
    expectTypeOf<Extract<keyof InteractionPolicy, 'batchDenied'>>().toEqualTypeOf<never>();
  });
});
