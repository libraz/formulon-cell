import {
  accentsFromTheme,
  CHART_PALETTE_PRESETS,
  type ChartPalettePreset,
  chartPalettePreset,
} from '../../charts/series-colors.js';
import {
  applySessionChartColors,
  applySessionChartElement,
  CHART_ELEMENT_ACTIONS,
  type ChartElementAction,
  chartElementState,
  chartUsesPalette,
  selectedSessionChart,
} from '../../commands/chart-design.js';
import type { SpreadsheetInstance } from '../../index.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';
import { setMenuControlDisabled } from './menu-feedback.js';

type ChartDesignDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyChartElementAction'
  | 'updateChartAddElementMenu'
  | 'applyChartColorsAction'
  | 'updateChartColorsMenu'
>;

const isChartElementAction = (value: string): value is ChartElementAction =>
  (CHART_ELEMENT_ACTIONS as readonly string[]).includes(value);

const isPalettePreset = (value: string): value is ChartPalettePreset =>
  Object.hasOwn(CHART_PALETTE_PRESETS, value);

const setChecked = (button: HTMLButtonElement, checked: boolean): void => {
  button.setAttribute('aria-checked', String(checked));
  button.classList.toggle('fc-tb__menu-item--active', checked);
};

export function createChartDesignDropdownDefaults(
  instance: SpreadsheetInstance,
): ChartDesignDropdownDefaults {
  const selectedChart = () => selectedSessionChart(instance.store.getState());

  return {
    applyChartElementAction(action) {
      if (action === 'more') {
        instance.openChartSettings('elements');
        return;
      }
      const chart = selectedChart();
      if (!chart || !isChartElementAction(action)) return;
      applySessionChartElement(
        instance.store,
        chart.id,
        action,
        instance.history,
        instance.workbook,
      );
    },
    updateChartAddElementMenu(menu) {
      const chart = selectedChart();
      for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-chart-element]')) {
        const action = button.dataset.chartElement ?? '';
        if (!isChartElementAction(action)) {
          setMenuControlDisabled(button, !chart);
          continue;
        }
        const state = chart ? chartElementState(chart, action) : null;
        setMenuControlDisabled(button, !state?.available);
        setChecked(button, state?.checked === true);
      }
    },
    applyChartColorsAction(preset) {
      const chart = selectedChart();
      if (!chart || !isPalettePreset(preset)) return;
      applySessionChartColors(
        instance.store,
        chart.id,
        preset,
        instance.history,
        instance.workbook,
      );
    },
    updateChartColorsMenu(menu) {
      const chart = selectedChart();
      const accents = accentsFromTheme(instance.workbook.getTheme()?.colors);
      for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-chart-colors]')) {
        const preset = button.dataset.chartColors ?? '';
        if (!isPalettePreset(preset)) continue;
        // The colourful palette is the workbook theme's.
        const colors = chartPalettePreset(preset, accents);
        button
          .querySelectorAll<HTMLElement>('.fc-tb__chart-palette-swatch')
          .forEach((swatch, i) => {
            swatch.style.backgroundColor = colors[i] ?? '';
          });
        setMenuControlDisabled(button, !chart || chart.kind === 'unsupported');
        setChecked(button, chart !== null && chartUsesPalette(chart, preset, instance.workbook));
      }
    },
  };
}
