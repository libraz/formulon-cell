import type {
  CellValue,
  FeatureFlags,
  FeatureId,
  SpreadsheetInstance,
  ThemeName,
} from '@libraz/formulon-cell';
import { useSelection } from '@libraz/formulon-cell-react';
import { type ReactElement, useCallback, useMemo, useState } from 'react';
import {
  type ChangeLogEntry,
  type DemoCommandStrings,
  type DemoUiStrings,
  demoSelectionLabel,
  evaluateDemoProbe,
  FEATURE_GROUPS,
  isDemoFeatureOn,
  LOCALES,
  PRESETS,
  type PresetKey,
  THEMES,
} from '../../demo-shared/index.js';

interface DemoOptionsPanelProps {
  ui: DemoUiStrings;
  commandText: DemoCommandStrings;
  instance: SpreadsheetInstance | null;
  hidden: boolean;
  theme: ThemeName;
  locale: string;
  preset: PresetKey;
  features: FeatureFlags;
  ribbon: boolean;
  formatters: { uppercase: boolean; arrows: boolean };
  log: ChangeLogEntry[];
  onThemeChange: (theme: ThemeName) => void;
  onLocaleChange: (locale: string) => void;
  onPresetChange: (preset: PresetKey) => void;
  onFeatureToggle: (id: FeatureId) => void;
  onRibbonChange: (checked: boolean) => void;
  onFormatterChange: (key: 'uppercase' | 'arrows', checked: boolean) => void;
}

export const DemoOptionsPanel = ({
  ui,
  commandText,
  instance,
  hidden,
  theme,
  locale,
  preset,
  features,
  ribbon,
  formatters,
  log,
  onThemeChange,
  onLocaleChange,
  onPresetChange,
  onFeatureToggle,
  onRibbonChange,
  onFormatterChange,
}: DemoOptionsPanelProps): ReactElement => {
  const [probe, setProbe] = useState<{ name: string; result: string } | null>(null);
  const selection = useSelection(instance);
  const selectionLabel = useMemo(() => demoSelectionLabel(selection), [selection]);

  const runProbe = useCallback(
    (name: string, args: CellValue[]) => {
      if (!instance) return;
      setProbe(evaluateDemoProbe(instance, name, args));
    },
    [instance],
  );

  return (
    <aside className="demo__panel" aria-label={ui.optionsPanel} hidden={hidden}>
      <section className="demo__card">
        <h2>{ui.demoChrome}</h2>
        <div className="demo__controls demo__controls--panel">
          <div className="demo__seg" role="group" aria-label={ui.theme}>
            {THEMES.map((t) => (
              <button
                key={t.value}
                type="button"
                className={`demo__seg-btn${t.value === theme ? ' demo__seg-btn--active' : ''}`}
                onClick={() => onThemeChange(t.value)}
                aria-pressed={t.value === theme}
              >
                {ui.themeLabels[t.value] ?? t.label}
              </button>
            ))}
          </div>
          <div className="demo__seg" role="group" aria-label={ui.locale}>
            {LOCALES.map((l) => (
              <button
                key={l.value}
                type="button"
                className={`demo__seg-btn${l.value === locale ? ' demo__seg-btn--active' : ''}`}
                onClick={() => onLocaleChange(l.value)}
                aria-pressed={l.value === locale}
              >
                {l.label}
              </button>
            ))}
          </div>
        </div>
      </section>

      <section className="demo__card">
        <h2>{ui.preset}</h2>
        <p className="demo__hint">{ui.presetHint}</p>
        <div className="demo__preset">
          {PRESETS.map((p) => (
            <button
              key={p.value}
              type="button"
              className={`demo__preset-btn${p.value === preset ? ' demo__preset-btn--active' : ''}`}
              onClick={() => onPresetChange(p.value)}
              aria-pressed={p.value === preset}
            >
              <span className="demo__preset-name">{ui.presets[p.value]?.label ?? p.label}</span>
              <span className="demo__preset-hint">{ui.presets[p.value]?.hint ?? p.hint}</span>
            </button>
          ))}
        </div>
      </section>

      <section className="demo__card">
        <h2>{ui.features}</h2>
        <p className="demo__hint">{ui.featuresHint}</p>
        {FEATURE_GROUPS.map((group) => (
          <div key={group.title} className="demo__feat-group">
            <h3 className="demo__feat-title">
              {ui.featureGroupLabels[group.title] ?? group.title}
            </h3>
            <div className="demo__feat-grid">
              {group.features.map((f) => {
                const enabled = isDemoFeatureOn(features, f.id);
                return (
                  <label key={f.id} className={`demo__feat${enabled ? ' demo__feat--on' : ''}`}>
                    <input
                      type="checkbox"
                      checked={enabled}
                      onChange={() => onFeatureToggle(f.id)}
                    />
                    <span>{ui.featureLabels[f.id] ?? f.label}</span>
                  </label>
                );
              })}
              {group.title === 'Chrome' ? (
                <label className={`demo__feat${ribbon ? ' demo__feat--on' : ''}`}>
                  <input
                    type="checkbox"
                    checked={ribbon}
                    onChange={(e) => onRibbonChange(e.target.checked)}
                  />
                  <span>{ui.spreadsheetRibbon}</span>
                </label>
              ) : null}
            </div>
          </div>
        ))}
      </section>

      <section className="demo__card">
        <h2>{commandText.selection}</h2>
        <p className="demo__mono">{selectionLabel}</p>
      </section>

      <section className="demo__card">
        <h2>{ui.cellRenderers}</h2>
        <p className="demo__hint">{ui.cellRenderersHint}</p>
        <label className="fc-tb__check">
          <input
            type="checkbox"
            checked={formatters.uppercase}
            onChange={(e) => onFormatterChange('uppercase', e.target.checked)}
          />
          {ui.uppercaseColumnA}
        </label>
        <label className="fc-tb__check">
          <input
            type="checkbox"
            checked={formatters.arrows}
            onChange={(e) => onFormatterChange('arrows', e.target.checked)}
          />
          {ui.arrowPrefixNegatives}
        </label>
      </section>

      <section className="demo__card">
        <h2>{ui.customFunctions}</h2>
        <p className="demo__hint">{ui.customFunctionsHint}</p>
        <div className="demo__probe">
          <button
            type="button"
            className="fc-tb__btn fc-tb__btn--ghost"
            onClick={() => runProbe('GREET', [{ kind: 'text', value: 'Workbook' }])}
            disabled={!instance}
          >
            GREET("Workbook")
          </button>
          <button
            type="button"
            className="fc-tb__btn fc-tb__btn--ghost"
            onClick={() => runProbe('FAHRENHEIT', [{ kind: 'number', value: 100 }])}
            disabled={!instance}
          >
            FAHRENHEIT(100)
          </button>
          {probe ? (
            <p className="demo__probe-out">
              → <code>{probe.result}</code>
            </p>
          ) : null}
        </div>
      </section>

      <section className="demo__card demo__card--log">
        <h2>{ui.cellChangeLog}</h2>
        <p className="demo__hint">{ui.cellChangeLogHint}</p>
        {log.length === 0 ? (
          <p className="fc-tb__empty">{ui.editCellToSeeEvents}</p>
        ) : (
          <ul className="demo__log">
            {log.map((entry) => (
              <li key={entry.id}>
                <span className="demo__log-cell">{entry.cell}</span>
                <span className="demo__log-arrow">→</span>
                <span className="demo__mono">{entry.preview}</span>
              </li>
            ))}
          </ul>
        )}
      </section>
    </aside>
  );
};
