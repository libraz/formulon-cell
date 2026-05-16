// Wide split-button used by the Cells group (Insert / Delete / Format / Fill /
// Clear / Sort / Find). The button renders an icon + label, the dropdown lists
// a polymorphic mix of regular items, section headings, radio-checked items,
// and optional pre-separator. Generic over `T` so callers pin a concrete
// action union without us re-typing it per group.

import { ChevronDown12Regular } from '@fluentui/react-icons';
import { Fragment, type ReactElement } from 'react';

import { Icon, type IconName } from '../icons.js';
import { RIBBON_KEYSHORTCUTS } from '../model.js';
import { useMenuOpen } from './useMenuOpen.js';

export interface CellMenuProps<T extends string> {
  command: string;
  disabled: boolean;
  icon: IconName;
  label: string;
  options: readonly {
    action: T;
    label: string;
    separatorBefore?: boolean;
    section?: boolean;
    active?: boolean;
  }[];
  activeAction?: T | null;
  activeButton?: boolean;
  onPick: (action: T) => void;
}

export function CellMenu<T extends string>({
  command,
  disabled,
  icon,
  label,
  options,
  activeAction,
  activeButton,
  onPick,
}: CellMenuProps<T>): ReactElement {
  const { open, setOpen, wrapRef } = useMenuOpen();

  return (
    <div
      ref={wrapRef}
      className={`demo__rb-menu${open ? ' demo__rb-menu--open' : ''}`}
      data-ribbon-command={command}
    >
      <button
        type="button"
        className={`demo__rb demo__rb-menu__btn demo__rb--wide${activeButton ? ' demo__rb--active' : ''}`}
        title={label}
        aria-label={label}
        aria-keyshortcuts={RIBBON_KEYSHORTCUTS[command]}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((next) => !next)}
      >
        <Icon name={icon} />
        <span>{label}</span>
        <ChevronDown12Regular className="demo__rb-menu__chev" aria-hidden="true" />
      </button>
      {open ? (
        <div className="demo__merge-menu demo__cell-menu" role="menu" aria-label={label}>
          {options.map((option) => {
            if (option.section) {
              return (
                <div
                  key={option.action}
                  className="demo__cf-menu__panel-title demo__cell-menu__section"
                  role="presentation"
                >
                  {option.label}
                </div>
              );
            }
            const checked = activeAction === option.action || option.active === true;
            const className = `demo__merge-menu__item${checked ? ' demo__rb--active' : ''}`;
            const radioLike = activeAction != null || option.active === true;
            const onClick = (): void => {
              onPick(option.action);
              setOpen(false);
            };
            return (
              <Fragment key={option.action}>
                {option.separatorBefore ? (
                  <div className="demo__cf-menu__sep" role="presentation" />
                ) : null}
                {!radioLike ? (
                  <button
                    type="button"
                    className={className}
                    role="menuitem"
                    data-cell-action={option.action}
                    onClick={onClick}
                  >
                    <Icon name={icon} />
                    <span>{option.label}</span>
                  </button>
                ) : (
                  <button
                    type="button"
                    className={className}
                    role="menuitemradio"
                    aria-checked={checked}
                    data-cell-action={option.action}
                    onClick={onClick}
                  >
                    <Icon name={icon} />
                    <span>{option.label}</span>
                  </button>
                )}
              </Fragment>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}
