import { SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { OperationIntent } from '../../../src/commands/interaction-policy.js';
import { defaultSparklineGroupSettings, setSparklines } from '../../../src/commands/sparkline.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { mountStubSheet } from '../../test-utils/mount.js';

// Instance-API writes authorize under a feature-level command id, as their siblings do
// (cellStyles, tableName, sparklineClear, slicerDelete), never under a ribbon button's id.
describe('instance-API drawing-object writes', () => {
  it('authorizes adding sparklines as sparklineInsert', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => wb, history });
    const off = registerInteractionController(store, controller);
    const seen: OperationIntent[] = [];
    controller.setPolicy({
      defaultOperation: 'allow',
      restrict: ({ intent }) => {
        seen.push(intent);
        return true;
      },
    });
    setSparklines(
      store,
      [{ addr: { sheet: 0, row: 0, col: 3 }, source: 'A1:C1' }],
      defaultSparklineGroupSettings(SparklineType.Line),
      history,
    );
    expect(seen.map((intent) => [intent.origin, intent.commandId])).toContainEqual([
      'instanceApi',
      'sparklineInsert',
    ]);
    off();
    controller.dispose();
    wb.dispose();
  });

  it('authorizes adding a slicer as slicerInsert', async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (const [row, cols] of [
      ['Region', 'Qty'],
      ['East', '1'],
      ['West', '2'],
    ].entries()) {
      for (const [col, text] of cols.entries()) workbook.setText({ sheet: 0, row, col }, text);
    }
    const sheet = await mountStubSheet({ locale: 'en', workbook, features: { slicer: true } });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    formatAsTable(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, { workbook });
    const tableName = sheet.instance.store.getState().tables.tables[0]?.name ?? '';
    const seen: OperationIntent[] = [];
    sheet.instance.commands.setPolicy({
      defaultOperation: 'allow',
      restrict: ({ intent }) => {
        seen.push(intent);
        return true;
      },
    });
    sheet.instance.addSlicer({ tableName, column: 'Region' });
    expect(seen.map((intent) => [intent.origin, intent.commandId])).toContainEqual([
      'instanceApi',
      'slicerInsert',
    ]);
    sheet.dispose();
  });
});
