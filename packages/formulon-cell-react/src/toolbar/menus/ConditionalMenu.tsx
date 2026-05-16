// Conditional Formatting split button with nested submenus for highlight /
// top-bottom / data bars / color scales / icon sets / clear, plus flat
// "new rule" and "manage". Swatch / icon-swatch buttons fire ConditionalMenuAction
// strings the playground/wrapper dispatcher resolves through handleConditionalAction.

import { ChevronDown12Regular } from '@fluentui/react-icons';
import {
  type ConditionalIconSetAction,
  type ConditionalMenuAction,
  type ConditionalPresetAction,
  conditionalColorScaleLabel,
  conditionalDataBarLabel,
  conditionalIconSetLabel,
  handleConditionalAction,
  type SpreadsheetInstance,
  type Strings,
} from '@libraz/formulon-cell';
import type { ReactElement } from 'react';

import { Icon } from '../icons.js';
import { useMenuOpen } from './useMenuOpen.js';

export interface ConditionalMenuProps {
  disabled: boolean;
  active: boolean;
  instance: SpreadsheetInstance | null;
  strings: Strings;
}

export function ConditionalMenu({
  disabled,
  active,
  instance,
  strings,
}: ConditionalMenuProps): ReactElement {
  const { open, setOpen, wrapRef } = useMenuOpen();
  const labels = strings.conditionalMenu;
  const dataBarLabel = (action: ConditionalPresetAction): string =>
    conditionalDataBarLabel(action, labels);
  const colorScaleLabel = (action: ConditionalPresetAction): string =>
    conditionalColorScaleLabel(action, labels);
  const iconSetLabel = (action: ConditionalIconSetAction): string =>
    conditionalIconSetLabel(action, labels);

  const onPick = (action: ConditionalMenuAction): void => {
    handleConditionalAction(instance, action);
    setOpen(false);
  };

  const item = (
    action: ConditionalMenuAction,
    label: string,
    key: string = action,
  ): ReactElement => (
    <button
      key={key}
      type="button"
      className="demo__merge-menu__item demo__cf-menu__item"
      role="menuitem"
      data-cf-action={action}
      onClick={() => onPick(action)}
    >
      <Icon name="conditional" />
      <span>{label}</span>
    </button>
  );

  const swatch = (
    action: ConditionalPresetAction,
    colors: readonly string[],
    label: string,
  ): ReactElement => {
    const colorCounts = new Map<string, number>();
    const swatchParts = colors.map((color) => {
      const count = (colorCounts.get(color) ?? 0) + 1;
      colorCounts.set(color, count);
      return { color, key: `${action}-${color}-${count}` };
    });
    return (
      <button
        key={action}
        type="button"
        className="demo__cf-menu__swatch"
        role="menuitem"
        data-cf-action={action}
        title={label}
        aria-label={label}
        onClick={() => onPick(action)}
      >
        {swatchParts.map((part) => (
          <span key={part.key} style={{ backgroundColor: part.color }} />
        ))}
      </button>
    );
  };

  const iconSwatch = (
    action: ConditionalIconSetAction,
    family: string,
    slots: readonly string[],
  ): ReactElement => (
    <button
      key={action}
      type="button"
      className="demo__cf-menu__iconset"
      role="menuitem"
      data-cf-action={action}
      title={iconSetLabel(action)}
      aria-label={iconSetLabel(action)}
      onClick={() => onPick(action)}
    >
      {slots.map((slot, index) => (
        <span
          key={`${action}-${slot}-${index}`}
          className={`demo__cf-icon demo__cf-icon--${family} demo__cf-icon--${slot}`}
        />
      ))}
    </button>
  );

  const iconSection = (label: string): ReactElement => (
    <div key={`section-${label}`} className="demo__cf-menu__panel-title" role="presentation">
      {label}
    </div>
  );

  const submenu = (label: string, children: ReactElement[], panelClass = ''): ReactElement => (
    <div className="demo__cf-menu__submenu" role="none">
      <button type="button" className="demo__merge-menu__item demo__cf-menu__item" role="menuitem">
        <Icon name="conditional" />
        <span>{label}</span>
        <span className="demo__cf-menu__arrow">›</span>
      </button>
      <div className={`demo__cf-menu__panel${panelClass ? ` ${panelClass}` : ''}`} role="menu">
        {children}
      </div>
    </div>
  );

  return (
    <div
      ref={wrapRef}
      className={`demo__rb-menu demo__cf-menu-wrap${open ? ' demo__rb-menu--open' : ''}`}
      data-ribbon-command="conditional"
    >
      <button
        type="button"
        className={`demo__rb demo__rb-menu__btn demo__rb--wide${active ? ' demo__rb--active' : ''}`}
        title={labels.title}
        aria-label={labels.title}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((next) => !next)}
      >
        <Icon name="conditional" />
        <span>{labels.title}</span>
        <ChevronDown12Regular className="demo__rb-menu__chev" aria-hidden="true" />
      </button>
      {open ? (
        <div className="demo__merge-menu demo__cf-menu" role="menu" aria-label={labels.title}>
          {submenu(labels.highlight, [
            item('cell-greater', labels.greater),
            item('cell-less', labels.less),
            item('cell-between', labels.between),
            item('cell-equal', labels.equal),
            item('text-contains', labels.textContains),
            item('date-occurring', labels.dateOccurring),
            item('duplicates', labels.duplicates),
            item('unique', labels.unique),
            item('highlight-more', labels.otherRules),
          ])}
          {submenu(labels.topBottom, [
            item('top10', labels.top10),
            item('bottom10', labels.bottom10),
            item('top10-percent', labels.top10Percent),
            item('bottom10-percent', labels.bottom10Percent),
            item('above-avg', labels.aboveAvg),
            item('below-avg', labels.belowAvg),
            item('top-bottom-more', labels.otherRules),
          ])}
          {submenu(labels.dataBars, [
            swatch('data-blue', ['#ffffff', '#638ec6'], dataBarLabel('data-blue')),
            swatch('data-green', ['#ffffff', '#63a95c'], dataBarLabel('data-green')),
            swatch('data-red', ['#ffffff', '#c45a5a'], dataBarLabel('data-red')),
            swatch('data-orange', ['#ffffff', '#d6a440'], dataBarLabel('data-orange')),
            swatch('data-purple', ['#ffffff', '#8a74b9'], dataBarLabel('data-purple')),
            swatch('data-teal', ['#ffffff', '#4ba1a8'], dataBarLabel('data-teal')),
            swatch('data-solid-blue', ['#4472c4', '#4472c4'], dataBarLabel('data-solid-blue')),
            swatch('data-solid-green', ['#70ad47', '#70ad47'], dataBarLabel('data-solid-green')),
            swatch('data-solid-red', ['#c00000', '#c00000'], dataBarLabel('data-solid-red')),
            swatch('data-solid-orange', ['#ed7d31', '#ed7d31'], dataBarLabel('data-solid-orange')),
            swatch('data-solid-purple', ['#8064a2', '#8064a2'], dataBarLabel('data-solid-purple')),
            swatch('data-solid-gray', ['#7f7f7f', '#7f7f7f'], dataBarLabel('data-solid-gray')),
            item('data-bars-more', labels.otherRules),
          ])}
          {submenu(labels.colorScales, [
            swatch('scale-gyr', ['#63be7b', '#ffeb84', '#f8696b'], colorScaleLabel('scale-gyr')),
            swatch('scale-ryg', ['#f8696b', '#ffeb84', '#63be7b'], colorScaleLabel('scale-ryg')),
            swatch('scale-gw', ['#63be7b', '#ffffff'], colorScaleLabel('scale-gw')),
            swatch('scale-rw', ['#f8696b', '#ffffff'], colorScaleLabel('scale-rw')),
            swatch('scale-bwr', ['#5a8dee', '#ffffff', '#f8696b'], colorScaleLabel('scale-bwr')),
            swatch('scale-rwb', ['#f8696b', '#ffffff', '#5a8dee'], colorScaleLabel('scale-rwb')),
            swatch('scale-gwg', ['#63be7b', '#ffffff', '#00a651'], colorScaleLabel('scale-gwg')),
            swatch('scale-ywg', ['#ffeb84', '#ffffff', '#63be7b'], colorScaleLabel('scale-ywg')),
            swatch('scale-rwr', ['#f8696b', '#ffffff', '#c00000'], colorScaleLabel('scale-rwr')),
            swatch('scale-bwb', ['#5a8dee', '#ffffff', '#4472c4'], colorScaleLabel('scale-bwb')),
            swatch('scale-yry', ['#ffeb84', '#f8696b', '#63be7b'], colorScaleLabel('scale-yry')),
            swatch('scale-gyg', ['#63be7b', '#ffeb84', '#00a651'], colorScaleLabel('scale-gyg')),
            item('color-scales-more', labels.otherRules),
          ])}
          {submenu(
            labels.iconSets,
            [
              iconSection(labels.direction),
              iconSwatch('icons-arrows3', 'arrow', ['up-green', 'right-yellow', 'down-red']),
              iconSwatch('icons-arrows5', 'arrow', [
                'up-green',
                'up-right-gray',
                'right-gray',
                'down-right-gray',
                'down-gray',
              ]),
              iconSwatch('icons-triangles3', 'triangle', ['up-green', 'flat-yellow', 'down-red']),
              iconSection(labels.shapes),
              iconSwatch('icons-traffic3', 'circle', ['green', 'yellow', 'red']),
              iconSwatch('icons-trafficRim3', 'rim', ['green', 'yellow', 'red']),
              iconSwatch('icons-symbols3', 'symbol', ['check-green', 'bang-yellow', 'x-red']),
              iconSwatch('icons-flags3', 'flag', ['green', 'yellow', 'red']),
              iconSection(labels.ratings),
              iconSwatch('icons-stars3', 'star', ['gold', 'half', 'empty']),
              iconSwatch('icons-quarters5', 'quarter', ['q4', 'q3', 'q2', 'q1', 'q0']),
              iconSwatch('icons-ratings5', 'rating', ['r4', 'r3', 'r2', 'r1', 'r0']),
              iconSwatch('icons-bars5', 'bars', ['b4', 'b3', 'b2', 'b1', 'b0']),
              iconSwatch('icons-boxes5', 'boxes', ['b4', 'b3', 'b2', 'b1', 'b0']),
              item('icon-sets-more', labels.otherRules),
            ],
            'demo__cf-menu__panel--icons',
          )}
          <div className="demo__cf-menu__sep" role="presentation" />
          {item('new-rule', labels.newRule)}
          {submenu(labels.clear, [
            item('clear-selection', labels.clearSelection),
            item('clear-sheet', labels.clearSheet),
          ])}
          {item('manage', labels.manage)}
        </div>
      ) : null}
    </div>
  );
}
