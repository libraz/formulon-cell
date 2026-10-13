import { describe, expect, it } from 'vitest';
import { chartDataFromTable } from '../../../src/charts/data.js';

describe('chartDataFromTable', () => {
  it('keeps named columns and category labels without summing values', () => {
    const data = chartDataFromTable(
      [
        ['Quarter', 'North', 'South', 'West'],
        ['Q1', 1, 2, 3],
        ['Q2', 4, 5, 6],
      ],
      { kind: 'column', seriesDirection: 'columns' },
    );

    expect(data).toEqual({
      domain: 'category',
      categories: ['Q1', 'Q2'],
      series: [
        { id: 'series-0', name: 'North', values: [1, 4] },
        { id: 'series-1', name: 'South', values: [2, 5] },
        { id: 'series-2', name: 'West', values: [3, 6] },
      ],
    });
  });

  it('uses rows for a wider auto-oriented table and keeps one-dimensional selections one series', () => {
    const rowData = chartDataFromTable(
      [
        ['', 'Q1', 'Q2', 'Q3'],
        ['North', 1, 2, 3],
        ['South', 4, 5, 6],
      ],
      { kind: 'line', seriesDirection: 'auto' },
    );
    expect(rowData).toMatchObject({
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'North', values: [1, 2, 3] },
        { name: 'South', values: [4, 5, 6] },
      ],
    });

    const oneRow = chartDataFromTable([[1, 2, 3]], { kind: 'column', seriesDirection: 'auto' });
    expect(oneRow.domain).toBe('category');
    expect(oneRow.series).toHaveLength(1);
    if (oneRow.domain !== 'category') throw new Error('expected category data');
    expect(oneRow.series[0]?.values).toEqual([1, 2, 3]);
  });

  it('handles category-only and series-name-only headers in either direction', () => {
    const categoryHeader = chartDataFromTable(
      [
        ['Jan', 'Feb'],
        [10, 15],
      ],
      { kind: 'line', seriesDirection: 'rows' },
    );
    expect(categoryHeader).toMatchObject({
      categories: ['Jan', 'Feb'],
      series: [{ values: [10, 15] }],
    });

    const seriesHeader = chartDataFromTable([['Sales', 10, 15]], {
      kind: 'column',
      seriesDirection: 'rows',
    });
    expect(seriesHeader).toMatchObject({
      categories: ['1', '2'],
      series: [{ name: 'Sales', values: [10, 15] }],
    });

    const oneColumn = chartDataFromTable([['Sales'], [10], [15]], {
      kind: 'column',
      seriesDirection: 'columns',
    });
    expect(oneColumn).toMatchObject({
      categories: ['1', '2'],
      series: [{ name: 'Sales', values: [10, 15] }],
    });
  });

  it('honors explicit false header overrides for one-dimensional selections', () => {
    const column = chartDataFromTable([['Sales'], [10], [15]], {
      kind: 'column',
      seriesDirection: 'columns',
      headerRow: false,
      headerColumn: false,
    });
    expect(column).toEqual({
      domain: 'category',
      categories: ['1', '2', '3'],
      series: [{ id: 'series-0', values: [null, 10, 15] }],
    });

    const row = chartDataFromTable([['Sales', 10, 15]], {
      kind: 'column',
      seriesDirection: 'rows',
      headerRow: false,
      headerColumn: false,
    });
    expect(row).toEqual({
      domain: 'category',
      categories: ['1', '2', '3'],
      series: [{ id: 'series-0', values: [null, 10, 15] }],
    });
  });

  it('normalizes invalid value cells to null', () => {
    const data = chartDataFromTable(
      [
        ['A', 'B'],
        ['x', 'bad'],
        ['y', Number.NaN],
      ],
      { kind: 'column', seriesDirection: 'columns' },
    );
    expect(data.domain).toBe('category');
    if (data.domain !== 'category') throw new Error('expected category data');
    expect(data.series[0]?.values).toEqual([null, null]);
    expect(data.series).toHaveLength(1);
  });

  it('preserves numeric scatter x coordinates and names y series', () => {
    const data = chartDataFromTable(
      [
        ['x', 'A', 'B'],
        [1, 2, 3],
        [10, 20, 30],
        [100, 200, 300],
      ],
      { kind: 'scatter', seriesDirection: 'columns' },
    );
    expect(data).toEqual({
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'A',
          points: [
            { x: 1, y: 2 },
            { x: 10, y: 20 },
            { x: 100, y: 200 },
          ],
        },
        {
          id: 'series-1',
          name: 'B',
          points: [
            { x: 1, y: 3 },
            { x: 10, y: 30 },
            { x: 100, y: 300 },
          ],
        },
      ],
    });
  });

  it('uses explicit header flags and labels numeric series names', () => {
    const data = chartDataFromTable(
      [
        ['Category', 2025, 2026],
        ['A', 1, 2],
        ['B', 3, 4],
      ],
      {
        kind: 'column',
        seriesDirection: 'columns',
        headerRow: true,
        headerColumn: true,
      },
    );

    expect(data).toEqual({
      domain: 'category',
      categories: ['A', 'B'],
      series: [
        { id: 'series-0', name: '2025', values: [1, 3] },
        { id: 'series-1', name: '2026', values: [2, 4] },
      ],
    });
  });

  it('infers partial textual headers without dropping category-first data', () => {
    const partial = chartDataFromTable(
      [
        ['Sales', null],
        [10, 20],
        [15, 25],
      ],
      { kind: 'column', seriesDirection: 'columns' },
    );
    expect(partial).toEqual({
      domain: 'category',
      categories: ['1', '2'],
      series: [
        { id: 'series-0', name: 'Sales', values: [10, 15] },
        { id: 'series-1', values: [20, 25] },
      ],
    });

    const rowPartial = chartDataFromTable(
      [
        ['Jan', null],
        [10, 20],
      ],
      { kind: 'column', seriesDirection: 'rows' },
    );
    expect(rowPartial).toEqual({
      domain: 'category',
      categories: ['Jan', '2'],
      series: [{ id: 'series-0', values: [10, 20] }],
    });

    const categoryFirst = chartDataFromTable(
      [
        ['Jan', 10],
        ['Feb', 15],
      ],
      { kind: 'column', seriesDirection: 'columns' },
    );
    expect(categoryFirst).toEqual({
      domain: 'category',
      categories: ['Jan', 'Feb'],
      series: [{ id: 'series-0', values: [10, 15] }],
    });
  });

  it('transposes explicit row-oriented scatter data and preserves blanks', () => {
    const data = chartDataFromTable(
      [
        ['X', 1, 10, 100],
        ['Y', 2, 20, 200],
      ],
      { kind: 'scatter', seriesDirection: 'rows' },
    );
    expect(data).toEqual({
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'Y',
          points: [
            { x: 1, y: 2 },
            { x: 10, y: 20 },
            { x: 100, y: 200 },
          ],
        },
      ],
    });

    const withGaps = chartDataFromTable(
      [
        ['x', 'Y'],
        [1, 2],
        [null, 3],
        [4, Number.NaN],
      ],
      { kind: 'scatter', seriesDirection: 'columns' },
    );
    expect(withGaps).toEqual({
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'Y',
          points: [
            { x: 1, y: 2 },
            { x: null, y: 3 },
            { x: 4, y: null },
          ],
        },
      ],
    });
  });

  it('maps row scatter header columns to transposed numeric series headers', () => {
    const data = chartDataFromTable(
      [
        [2025, 1, 10],
        [2026, 2, 20],
      ],
      { kind: 'scatter', seriesDirection: 'rows', headerColumn: true },
    );
    expect(data).toEqual({
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: '2026',
          points: [
            { x: 1, y: 2 },
            { x: 10, y: 20 },
          ],
        },
      ],
    });
  });

  it('keeps a category-first partial header row when later category labels are textual', () => {
    const data = chartDataFromTable(
      [
        ['Jan', null],
        ['Feb', 10],
        ['Mar', 20],
      ],
      { kind: 'column', seriesDirection: 'columns' },
    );
    expect(data).toEqual({
      domain: 'category',
      categories: ['Jan', 'Feb', 'Mar'],
      series: [{ id: 'series-0', values: [null, 10, 20] }],
    });
  });
});
