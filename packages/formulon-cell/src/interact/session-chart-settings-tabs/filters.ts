/** Filters tab of the chart settings dialog: series and category visibility. */

import type { ChartData } from '../../charts/types.js';
import { element, makeCheckbox, textNode } from '../session-chart-settings-dom.js';
import {
  DEFAULT_LABELS,
  type SessionChartSettingsLabels,
} from '../session-chart-settings-labels.js';
import { type ChartDraft, seriesEntries } from '../session-chart-settings-model.js';

const categoryEntries = (
  data: ChartData,
  labels?: SessionChartSettingsLabels,
): readonly { id: string; label: string }[] => {
  if (data.domain === 'category') {
    return data.categories.map((category, index) => ({
      id: data.categoryIds?.[index] ?? `category-${index}`,
      label: category || `${labels?.categories ?? DEFAULT_LABELS.categories} ${index + 1}`,
    }));
  }
  const count = data.series.reduce((max, series) => Math.max(max, series.points.length), 0);
  return Array.from({ length: count }, (_, index) => ({
    id: `point-${index}`,
    label: `${labels?.point ?? DEFAULT_LABELS.point} ${index + 1}`,
  }));
};

export const renderFilters = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  data: ChartData,
  draft: ChartDraft,
): void => {
  const series = element('fieldset', 'fc-chart-settings__group');
  series.appendChild(textNode('legend', labels.series, 'fc-chart-settings__legend'));
  const seriesEntriesValue = seriesEntries(data, labels);
  if (seriesEntriesValue.length === 0) {
    series.appendChild(textNode('div', labels.noSeries, 'fc-chart-settings__empty'));
  } else {
    const actions = element('div', 'fc-chart-settings__filter-actions');
    const all = element('button');
    all.type = 'button';
    all.textContent = labels.selectAll;
    all.addEventListener('click', () => {
      const known = new Set(seriesEntriesValue.map((entry) => entry.id));
      draft.filter.hiddenSeriesIds = draft.filter.hiddenSeriesIds.filter((id) => !known.has(id));
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    const clear = element('button');
    clear.type = 'button';
    clear.textContent = labels.clearAll;
    clear.addEventListener('click', () => {
      draft.filter.hiddenSeriesIds = [
        ...new Set([
          ...draft.filter.hiddenSeriesIds,
          ...seriesEntriesValue.map((entry) => entry.id),
        ]),
      ];
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    actions.append(all, clear);
    series.appendChild(actions);
    for (const entry of seriesEntriesValue) {
      makeCheckbox(
        series,
        entry.label,
        !draft.filter.hiddenSeriesIds.includes(entry.id),
        (checked) => {
          draft.filter.hiddenSeriesIds = checked
            ? draft.filter.hiddenSeriesIds.filter((id) => id !== entry.id)
            : [...new Set([...draft.filter.hiddenSeriesIds, entry.id])];
        },
        false,
        'fc-chart-settings__filter-item',
      );
    }
  }
  body.appendChild(series);

  const categories = element('fieldset', 'fc-chart-settings__group');
  categories.appendChild(textNode('legend', labels.categories, 'fc-chart-settings__legend'));
  const categoryEntriesValue = categoryEntries(data, labels);
  if (categoryEntriesValue.length === 0) {
    categories.appendChild(textNode('div', labels.noCategories, 'fc-chart-settings__empty'));
  } else {
    const actions = element('div', 'fc-chart-settings__filter-actions');
    const all = element('button');
    all.type = 'button';
    all.textContent = labels.selectAll;
    all.addEventListener('click', () => {
      const known = new Set(categoryEntriesValue.map((entry) => entry.id));
      draft.filter.hiddenCategoryIds = draft.filter.hiddenCategoryIds.filter(
        (id) => !known.has(id),
      );
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    const clear = element('button');
    clear.type = 'button';
    clear.textContent = labels.clearAll;
    clear.addEventListener('click', () => {
      draft.filter.hiddenCategoryIds = [
        ...new Set([
          ...draft.filter.hiddenCategoryIds,
          ...categoryEntriesValue.map((entry) => entry.id),
        ]),
      ];
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    actions.append(all, clear);
    categories.appendChild(actions);
    for (const entry of categoryEntriesValue) {
      makeCheckbox(
        categories,
        entry.label,
        !draft.filter.hiddenCategoryIds.includes(entry.id),
        (checked) => {
          draft.filter.hiddenCategoryIds = checked
            ? draft.filter.hiddenCategoryIds.filter((id) => id !== entry.id)
            : [...new Set([...draft.filter.hiddenCategoryIds, entry.id])];
        },
        false,
        'fc-chart-settings__filter-item',
      );
    }
  }
  body.appendChild(categories);
};
