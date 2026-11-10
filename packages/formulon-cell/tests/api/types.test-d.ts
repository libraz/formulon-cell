import type { Value } from '@libraz/formulon';
import type {
  CommandId,
  DialogId,
  Disposer,
  SaveResult,
  Sheet,
  SpreadsheetFailure,
  SpreadsheetInstance,
  Workbook,
} from '@libraz/formulon-cell';
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
});
