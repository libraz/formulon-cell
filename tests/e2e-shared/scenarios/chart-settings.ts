import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runChartSettingsScenario(page: Page): Promise<void> {
  const sheet = new SpreadsheetPage(page);
  await sheet.mount();
  await sheet.expectNoStub();
  const commandsUrl = `/@fs${new URL('../../../packages/formulon-cell/src/commands/session-chart.ts', import.meta.url).pathname}`;
  await page.evaluate(
    async ({ commandsUrl }) => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const { createSessionChart } = await import(commandsUrl);
      const table = [
        ['Month', 'Sales', 'Costs'],
        ['Jan', 10, 6],
        ['Feb', 15, 8],
        ['Mar', 8, 3],
      ];
      const result = instance.applyChanges(
        table.flatMap((row, r) =>
          row.map((value, col) => ({
            addr: { sheet: 0, row: r, col },
            value:
              typeof value === 'number'
                ? { kind: 'number' as const, value }
                : { kind: 'text' as const, value },
          })),
        ),
      );
      if (result.status !== 'applied') throw new Error('chart fixture cells were not applied');
      createSessionChart(
        instance.store,
        { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
        {
          id: 'settings-e2e',
          kind: 'column',
          seriesDirection: 'columns',
          title: 'Monthly sales',
          x: 140,
          y: 60,
          w: 420,
          h: 280,
        },
        instance.history,
      );
    },
    { commandsUrl },
  );

  const chart = page.locator('[data-chart-id="settings-e2e"]');
  const bars = chart.locator('[data-chart-role="bar"]');
  await expect(bars).toHaveCount(6);
  const originalFills = await bars.evaluateAll((nodes) =>
    nodes.map((node) => node.getAttribute('fill')),
  );
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  const dialog = page.getByRole('dialog', { name: 'Chart settings', exact: true });
  await dialog.getByLabel('Show data labels', { exact: true }).check();
  await dialog.getByLabel('Category name', { exact: true }).check();
  await dialog
    .getByRole('group', { name: 'Data labels', exact: true })
    .getByRole('textbox', { name: 'Number format', exact: true })
    .fill('#,##0.00');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(chart.locator('[data-chart-role="data-label"]')).toHaveCount(6);
  await expect(chart.locator('[data-chart-role="data-label"]').first()).toHaveText('Jan, 10.00');
  expect(await bars.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('fill')))).toEqual(
    originalFills,
  );

  await chart.locator('[data-chart-settings-tab="filters"]').click();
  await expect(dialog.getByRole('tab', { name: 'Filters', exact: true })).toBeFocused();
  await dialog.getByLabel('Costs', { exact: true }).uncheck();
  await dialog.getByLabel('Feb', { exact: true }).uncheck();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(bars).toHaveCount(2);
  expect(
    await bars.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('data-point-id'))),
  ).toEqual(['category-0', 'category-2']);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  await expect(bars).toHaveCount(6);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(bars).toHaveCount(2);

  await chart.locator('[data-chart-settings-tab="styles"]').click();
  await dialog.getByRole('combobox', { name: 'Point', exact: true }).selectOption('category-2');
  const pointFill = dialog
    .getByRole('textbox', { name: 'Fill color', exact: true })
    .and(dialog.locator('input[type="text"]'));
  await pointFill.fill('#ff0000');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const styledPoint = chart.locator('[data-chart-role="bar"][data-point-id="category-2"]');
  await expect(styledPoint).toHaveAttribute('fill', '#ff0000');
  await expect(
    chart.locator('[data-chart-role="bar"][data-point-id="category-0"]'),
  ).toHaveAttribute('fill', originalFills[0] ?? '');

  await chart.locator('[data-chart-settings-tab="styles"]').click();
  await dialog.getByRole('combobox', { name: 'Point', exact: true }).selectOption('category-2');
  await pointFill.fill('#00ff00');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(styledPoint).toHaveAttribute('fill', '#ff0000');
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  await expect(styledPoint).toHaveAttribute('fill', originalFills[0] ?? '');
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(styledPoint).toHaveAttribute('fill', '#ff0000');

  await chart.locator('[data-chart-settings-tab="styles"]').click();
  await dialog.getByRole('combobox', { name: 'Point', exact: true }).selectOption('category-0');
  await dialog
    .getByRole('spinbutton', { name: 'Fill color Transparency (%)', exact: true })
    .fill('25');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const inheritedPoint = chart.locator('[data-chart-role="bar"][data-point-id="category-0"]');
  await expect(inheritedPoint).toHaveAttribute('fill-opacity', '0.75');
  await chart.locator('[data-chart-settings-tab="styles"]').click();
  await dialog
    .getByRole('combobox', { name: 'Point', exact: true })
    .selectOption({ label: 'All points' });
  await pointFill.fill('#00aa00');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(inheritedPoint).toHaveAttribute('fill', '#00aa00');
  await expect(inheritedPoint).toHaveAttribute('fill-opacity', '0.75');
  await expect(styledPoint).toHaveAttribute('fill', '#ff0000');

  const automaticHeight = Number(await inheritedPoint.getAttribute('height'));
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  const yAxis = dialog.getByRole('group', { name: 'Y-axis', exact: true });
  await yAxis.getByRole('spinbutton', { name: 'Minimum', exact: true }).fill('0');
  await yAxis.getByRole('spinbutton', { name: 'Maximum', exact: true }).fill('20');
  await yAxis.getByRole('spinbutton', { name: 'Major unit', exact: true }).fill('5');
  await yAxis.getByRole('textbox', { name: 'Number format', exact: true }).fill('0.0');
  await yAxis.getByRole('checkbox', { name: 'Major gridlines', exact: true }).uncheck();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const yAxisLabels = chart.locator('[data-chart-role="axis-label"][data-chart-axis="y"]');
  await expect(yAxisLabels).toHaveText(['0.0', '5.0', '10.0', '15.0', '20.0']);
  await expect(chart.locator('[data-chart-role="axis-tick"][data-chart-axis="y"]')).toHaveCount(0);
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(automaticHeight / 2);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(automaticHeight);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(yAxisLabels).toHaveText(['0.0', '5.0', '10.0', '15.0', '20.0']);

  const xAxis = dialog.getByRole('group', { name: 'X-axis', exact: true });
  const beforeReverse = {
    x: Number(await inheritedPoint.getAttribute('x')),
    y: Number(await inheritedPoint.getAttribute('y')),
    width: Number(await inheritedPoint.getAttribute('width')),
    height: Number(await inheritedPoint.getAttribute('height')),
  };
  const otherCenter =
    Number(await styledPoint.getAttribute('x')) +
    Number(await styledPoint.getAttribute('width')) / 2;
  const plotTop = Number(await yAxisLabels.last().getAttribute('y')) - 3;
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('checkbox', { name: 'Reverse order', exact: true }).check();
  await yAxis.getByRole('checkbox', { name: 'Reverse order', exact: true }).check();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(
    Number(await inheritedPoint.getAttribute('x')) +
      Number(await inheritedPoint.getAttribute('width')) / 2,
  ).toBeCloseTo(otherCenter);
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(plotTop);
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(beforeReverse.height);
  await expect(styledPoint).toHaveAttribute('fill', '#ff0000');
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  expect(Number(await inheritedPoint.getAttribute('x'))).toBeCloseTo(beforeReverse.x);
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(beforeReverse.y);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(plotTop);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('checkbox', { name: 'Reverse order', exact: true }).uncheck();
  await yAxis.getByRole('checkbox', { name: 'Reverse order', exact: true }).uncheck();
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(plotTop);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('checkbox', { name: 'Reverse order', exact: true }).uncheck();
  await yAxis.getByRole('checkbox', { name: 'Reverse order', exact: true }).uncheck();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await inheritedPoint.getAttribute('x'))).toBeCloseTo(beforeReverse.x);
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(beforeReverse.y);

  const horizontalAxis = chart.locator('[data-chart-role="axis"][data-chart-axis="x"]');
  const verticalAxis = chart.locator('[data-chart-role="axis"][data-chart-axis="y"]');
  const automaticCrossing = {
    x: Number(await verticalAxis.getAttribute('x1')),
    y: Number(await horizontalAxis.getAttribute('y1')),
  };
  const plotRight = Number(await horizontalAxis.getAttribute('x2'));
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis
    .getByRole('combobox', { name: 'Vertical axis crosses', exact: true })
    .selectOption('max');
  await yAxis
    .getByRole('combobox', { name: 'Horizontal axis crosses', exact: true })
    .selectOption('custom');
  await yAxis.getByRole('spinbutton', { name: 'Custom value', exact: true }).fill('10');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await verticalAxis.getAttribute('x1'))).toBeCloseTo(plotRight);
  expect(Number(await horizontalAxis.getAttribute('y1'))).toBeCloseTo(
    (plotTop + automaticCrossing.y) / 2,
  );
  expect(Number(await inheritedPoint.getAttribute('x'))).toBeCloseTo(beforeReverse.x);
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(beforeReverse.y);
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(beforeReverse.height);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  expect(Number(await verticalAxis.getAttribute('x1'))).toBeCloseTo(automaticCrossing.x);
  expect(Number(await horizontalAxis.getAttribute('y1'))).toBeCloseTo(automaticCrossing.y);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  expect(Number(await verticalAxis.getAttribute('x1'))).toBeCloseTo(plotRight);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis
    .getByRole('combobox', { name: 'Vertical axis crosses', exact: true })
    .selectOption('min');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  expect(Number(await verticalAxis.getAttribute('x1'))).toBeCloseTo(plotRight);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis
    .getByRole('combobox', { name: 'Vertical axis crosses', exact: true })
    .selectOption('auto');
  await yAxis
    .getByRole('combobox', { name: 'Horizontal axis crosses', exact: true })
    .selectOption('auto');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await verticalAxis.getAttribute('x1'))).toBeCloseTo(automaticCrossing.x);
  expect(Number(await horizontalAxis.getAttribute('y1'))).toBeCloseTo(automaticCrossing.y);

  const xCategoryLabels = chart.locator('[data-chart-role="category-label"][data-chart-axis="x"]');
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('high');
  await yAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('high');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await xCategoryLabels.first().getAttribute('y'))).toBeCloseTo(
    Number(await verticalAxis.getAttribute('y1')) - 4,
  );
  expect(Number(await yAxisLabels.first().getAttribute('x'))).toBeCloseTo(
    Number(await horizontalAxis.getAttribute('x2')) + 4,
  );
  await expect(yAxisLabels.first()).toHaveAttribute('text-anchor', 'start');
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(beforeReverse.y);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(yAxisLabels.first()).toHaveAttribute('text-anchor', 'start');

  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('nextTo');
  await yAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('nextTo');
  await xAxis
    .getByRole('combobox', { name: 'Vertical axis crosses', exact: true })
    .selectOption('max');
  await yAxis
    .getByRole('combobox', { name: 'Horizontal axis crosses', exact: true })
    .selectOption('custom');
  await yAxis.getByRole('spinbutton', { name: 'Custom value', exact: true }).fill('10');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await xCategoryLabels.first().getAttribute('y'))).toBeCloseTo(
    Number(await horizontalAxis.getAttribute('y1')) + 12,
  );
  expect(Number(await yAxisLabels.first().getAttribute('x'))).toBeCloseTo(
    Number(await verticalAxis.getAttribute('x1')) - 2,
  );
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(beforeReverse.height);

  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('none');
  await yAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('none');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(xCategoryLabels).toHaveCount(0);
  await expect(yAxisLabels).toHaveCount(0);
  await expect(horizontalAxis).toHaveCount(1);
  await expect(verticalAxis).toHaveCount(1);
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(beforeReverse.height);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('combobox', { name: 'Label position', exact: true }).selectOption('low');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(xCategoryLabels).toHaveCount(0);

  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis
    .getByRole('combobox', { name: 'Label position', exact: true })
    .selectOption('automatic');
  await yAxis
    .getByRole('combobox', { name: 'Label position', exact: true })
    .selectOption('automatic');
  await xAxis
    .getByRole('combobox', { name: 'Vertical axis crosses', exact: true })
    .selectOption('auto');
  await yAxis
    .getByRole('combobox', { name: 'Horizontal axis crosses', exact: true })
    .selectOption('auto');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(xCategoryLabels).toHaveText(['Jan', 'Mar']);
  await expect(yAxisLabels).toHaveText(['0.0', '5.0', '10.0', '15.0', '20.0']);
  expect(Number(await inheritedPoint.getAttribute('x'))).toBeCloseTo(beforeReverse.x);
  expect(Number(await inheritedPoint.getAttribute('y'))).toBeCloseTo(beforeReverse.y);

  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await expect(
    xAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }),
  ).toBeDisabled();
  await yAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }).check();
  await expect(yAxis.getByRole('spinbutton', { name: 'Log base', exact: true })).toHaveValue('10');
  await expect(yAxis.getByRole('spinbutton', { name: 'Major unit', exact: true })).toBeDisabled();
  await yAxis.getByRole('spinbutton', { name: 'Minimum', exact: true }).fill('1');
  await yAxis.getByRole('spinbutton', { name: 'Maximum', exact: true }).fill('1000');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(yAxisLabels).toHaveText(['1.0', '10.0', '100.0', '1000.0']);
  const logPlotHeight =
    Number(await verticalAxis.getAttribute('y2')) - Number(await verticalAxis.getAttribute('y1'));
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(logPlotHeight / 3);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  await expect(yAxisLabels).toHaveText(['0.0', '5.0', '10.0', '15.0', '20.0']);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(yAxisLabels).toHaveText(['1.0', '10.0', '100.0', '1000.0']);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await yAxis.getByRole('spinbutton', { name: 'Log base', exact: true }).fill('2');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(yAxisLabels).toHaveText(['1.0', '10.0', '100.0', '1000.0']);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await yAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }).uncheck();
  await expect(yAxis.getByRole('spinbutton', { name: 'Major unit', exact: true })).toBeEnabled();
  await expect(yAxis.getByRole('spinbutton', { name: 'Major unit', exact: true })).toHaveValue('5');
  await yAxis.getByRole('spinbutton', { name: 'Minimum', exact: true }).fill('0');
  await yAxis.getByRole('spinbutton', { name: 'Maximum', exact: true }).fill('20');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await inheritedPoint.getAttribute('height'))).toBeCloseTo(beforeReverse.height);

  const automaticWidth = Number(await inheritedPoint.getAttribute('width'));
  await chart.locator('[data-chart-settings-tab="styles"]').click();
  await dialog.getByRole('spinbutton', { name: 'Gap width (%)', exact: true }).fill('0');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  expect(Number(await inheritedPoint.getAttribute('width'))).toBeGreaterThan(automaticWidth);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  expect(Number(await inheritedPoint.getAttribute('width'))).toBeCloseTo(automaticWidth);

  const layoutUrl = `/@fs${new URL('../../../packages/formulon-cell/src/charts/layout.ts', import.meta.url).pathname}`;
  const narrowScene = await page.evaluate(
    async ({ layoutUrl }) => {
      const { layoutChart } = await import(layoutUrl);
      const axis = { numberFormat: 'host-code', min: 0, max: 20, majorUnit: 5 };
      const scene = layoutChart({
        spec: { kind: 'scatter', xAxis: axis, yAxis: { ...axis, title: '値' } },
        data: {
          domain: 'xy',
          series: [
            {
              id: 's',
              points: [
                { x: 5, y: 10 },
                { x: 15, y: 20 },
              ],
            },
          ],
        },
        width: 240,
        height: 132,
        formatValue: (value: number) => `日本語の長い数値書式 ${value.toFixed(2)}円`,
      });
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', '0 0 240 132');
      svg.style.cssText = 'position:fixed;left:0;top:0;width:240px;height:132px';
      for (const primitive of scene.elements) {
        const node = document.createElementNS(svg.namespaceURI, primitive.tag);
        for (const [name, value] of Object.entries(primitive.attrs))
          node.setAttribute(name, String(value));
        if (primitive.text !== undefined) node.textContent = primitive.text;
        if (primitive.role) node.setAttribute('data-chart-role', primitive.role);
        svg.append(node);
      }
      document.body.append(svg);
      try {
        const bounds = svg.getBoundingClientRect();
        const clipped = [
          ...svg.querySelectorAll('[data-chart-role="axis-label"], [data-chart-role="axis-title"]'),
        ]
          .filter((node) => {
            const box = node.getBoundingClientRect();
            return (
              box.left < bounds.left ||
              box.right > bounds.right ||
              box.top < bounds.top ||
              box.bottom > bounds.bottom
            );
          })
          .map((node) => node.textContent);
        return { plotWidth: scene.plot.width, clipped };
      } finally {
        svg.remove();
      }
    },
    { layoutUrl },
  );
  expect(narrowScene.plotWidth).toBeGreaterThanOrEqual(48);
  expect(narrowScene.clipped).toEqual([]);

  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog.getByRole('combobox', { name: 'Chart type', exact: true }).selectOption('doughnut');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const slices = chart.locator('[data-chart-role="slice"]');
  await expect(slices).toHaveCount(2);
  await expect(bars).toHaveCount(0);
  await expect(chart.locator('[data-chart-role="axis"]')).toHaveCount(0);
  const originalSlice = await slices.first().getAttribute('d');
  await chart.locator('[data-chart-settings-tab="styles"]').click();
  const holeSize = dialog.getByRole('spinbutton', { name: 'Hole size (%)', exact: true });
  await expect(holeSize).toHaveValue('50');
  await holeSize.fill('70');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const resizedSlice = await slices.first().getAttribute('d');
  expect(resizedSlice).not.toBe(originalSlice);
  const radii = [...(resizedSlice ?? '').matchAll(/A\s+([-\d.e+]+)\s+/gi)].map((match) =>
    Number(match[1]),
  );
  const outerRadius = radii[0];
  const innerRadius = radii.at(-1);
  if (outerRadius === undefined || innerRadius === undefined || outerRadius <= 0) {
    throw new Error('Doughnut path must contain outer and inner arcs');
  }
  expect(innerRadius / outerRadius).toBeCloseTo(0.7);
  expect(
    await slices.evaluateAll((nodes) => {
      const bounds = nodes.map((node) => (node as SVGGraphicsElement).getBBox());
      const minX = Math.min(...bounds.map((box) => box.x));
      const maxX = Math.max(...bounds.map((box) => box.x + box.width));
      const minY = Math.min(...bounds.map((box) => box.y));
      const maxY = Math.max(...bounds.map((box) => box.y + box.height));
      const center = new DOMPoint((minX + maxX) / 2, (minY + maxY) / 2);
      return nodes.every((node) => !(node as SVGGeometryElement).isPointInFill(center));
    }),
  ).toBe(true);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  await expect(slices.first()).toHaveAttribute('d', originalSlice ?? '');
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(slices.first()).toHaveAttribute('d', resizedSlice ?? '');

  await chart.locator('[data-chart-settings-tab="filters"]').click();
  await dialog.getByLabel('Costs', { exact: true }).check();
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(slices).toHaveCount(4);
  expect(
    await slices.evaluateAll((nodes) => [
      ...new Set(nodes.map((node) => node.getAttribute('data-series-id'))),
    ]),
  ).toEqual(['series-0', 'series-1']);

  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog.getByRole('combobox', { name: 'Chart type', exact: true }).selectOption('scatter');
  await dialog.getByRole('tab', { name: 'Styles', exact: true }).click();
  await dialog.getByRole('combobox', { name: 'Point', exact: true }).selectOption('point-1');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(slices).toHaveCount(4);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog.getByRole('combobox', { name: 'Chart type', exact: true }).selectOption('scatter');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(3);
  await expect(slices).toHaveCount(0);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  await expect(slices).toHaveCount(4);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(3);
  const connectors = chart.locator('[data-chart-role="line"]');
  await expect(connectors).toHaveCount(0);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog.getByRole('combobox', { name: 'Scatter style', exact: true }).selectOption('line');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(connectors).toHaveCount(2);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(0);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
  );
  await expect(connectors).toHaveCount(0);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(3);
  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
  );
  await expect(connectors).toHaveCount(2);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(0);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog.getByRole('combobox', { name: 'Scatter style', exact: true }).selectOption('marker');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(connectors).toHaveCount(2);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(0);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog
    .getByRole('combobox', { name: 'Scatter style', exact: true })
    .selectOption('lineMarker');
  await dialog.getByRole('tab', { name: 'Styles', exact: true }).click();
  await dialog.getByRole('combobox', { name: 'Point', exact: true }).selectOption('__all__');
  await dialog
    .getByRole('textbox', { name: 'Line color', exact: true })
    .and(dialog.locator('input[type="text"]'))
    .fill('#225588');
  await dialog.getByRole('combobox', { name: 'Line dash', exact: true }).selectOption('dash');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(connectors).toHaveCount(2);
  await expect(connectors.first()).toHaveAttribute('stroke', '#225588');
  await expect(connectors.first()).toHaveAttribute('stroke-dasharray', '6 4');
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(3);
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const result = instance.applyChanges([{ addr: { sheet: 0, row: 2, col: 2 }, input: '' }]);
    if (result.status !== 'applied') throw new Error('Scatter blank fixture was not applied');
  });
  await expect(connectors).toHaveCount(0);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(2);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await dialog.getByRole('combobox', { name: 'Empty cells', exact: true }).selectOption('span');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(connectors).toHaveCount(1);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(2);
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const result = instance.applyChanges([{ addr: { sheet: 0, row: 2, col: 2 }, input: '8' }]);
    if (result.status !== 'applied') throw new Error('Scatter source fixture was not restored');
  });
  await expect(connectors).toHaveCount(2);
  await expect(chart.locator('[data-chart-role="point"]')).toHaveCount(3);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }).check();
  await yAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }).check();
  for (const axis of [xAxis, yAxis]) {
    await axis.getByRole('spinbutton', { name: 'Minimum', exact: true }).fill('1');
    await axis.getByRole('spinbutton', { name: 'Maximum', exact: true }).fill('100');
  }
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const logPoints = chart.locator('[data-chart-role="point"]');
  await expect(logPoints).toHaveCount(3);
  await expect(connectors).toHaveCount(2);
  expect(Number(await connectors.first().getAttribute('x1'))).toBeCloseTo(
    Number(await logPoints.first().getAttribute('cx')),
  );
  expect(Number(await connectors.first().getAttribute('y1'))).toBeCloseTo(
    Number(await logPoints.first().getAttribute('cy')),
  );
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const result = instance.applyChanges([{ addr: { sheet: 0, row: 2, col: 2 }, input: '0' }]);
    if (result.status !== 'applied') throw new Error('Log scatter zero fixture was not applied');
  });
  await expect(logPoints).toHaveCount(2);
  await expect(connectors).toHaveCount(0);
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const result = instance.applyChanges([{ addr: { sheet: 0, row: 2, col: 2 }, input: '8' }]);
    if (result.status !== 'applied') throw new Error('Log scatter source fixture was not restored');
  });
  await expect(connectors).toHaveCount(2);
  await chart.locator('[data-chart-settings-tab="elements"]').click();
  await xAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }).uncheck();
  await yAxis.getByRole('checkbox', { name: 'Logarithmic scale', exact: true }).uncheck();
  await yAxis.getByRole('spinbutton', { name: 'Minimum', exact: true }).fill('0');
  await yAxis.getByRole('spinbutton', { name: 'Maximum', exact: true }).fill('20');
  await dialog.getByRole('combobox', { name: 'Chart type', exact: true }).selectOption('column');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(bars).toHaveCount(4);
  await expect(
    chart.locator('[data-chart-role="bar"][data-series-id="series-0"][data-point-id="category-2"]'),
  ).toHaveAttribute('fill', '#ff0000');
}
