// Merge cells split button with a dropdown of the four merge variants. The
// active "center" / "merge" state mirrors what's painted on the underlying
// active cell anchor.

import { ChevronDown12Regular } from '@fluentui/react-icons';
import type { MergeAction } from '@libraz/formulon-cell';
import type { ReactElement } from 'react';

import { Icon } from '../icons.js';
import { useMenuOpen } from './useMenuOpen.js';

export interface MergeMenuProps {
  disabled: boolean;
  activeAction?: 'mergeCenter' | 'mergeCells' | null;
  labels: {
    mergeAndCenter: string;
    mergeAcross: string;
    mergeCells: string;
    unmergeCells: string;
  };
  onPick: (action: MergeAction) => void;
}

export function MergeMenu({
  disabled,
  activeAction,
  labels,
  onPick,
}: MergeMenuProps): ReactElement {
  const { open, setOpen, wrapRef } = useMenuOpen();
  const options: readonly { action: MergeAction; label: string }[] = [
    { action: 'mergeCenter', label: labels.mergeAndCenter },
    { action: 'mergeAcross', label: labels.mergeAcross },
    { action: 'mergeCells', label: labels.mergeCells },
    { action: 'unmergeCells', label: labels.unmergeCells },
  ];

  return (
    <div
      ref={wrapRef}
      className={`demo__rb-menu${open ? ' demo__rb-menu--open' : ''}`}
      data-ribbon-command="merge"
    >
      <button
        type="button"
        className={`demo__rb demo__rb-menu__btn${activeAction ? ' demo__rb--active' : ''}`}
        title={labels.mergeCells}
        aria-label={labels.mergeCells}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((next) => !next)}
      >
        <Icon name="merge" />
        <ChevronDown12Regular className="demo__rb-menu__chev" aria-hidden="true" />
      </button>
      {open ? (
        <div className="demo__merge-menu" role="menu" aria-label={labels.mergeCells}>
          {options.map((option) => {
            const checked = activeAction === option.action;
            return (
              <button
                key={option.action}
                type="button"
                className={`demo__merge-menu__item${checked ? ' demo__rb--active' : ''}`}
                role={
                  option.action === 'unmergeCells' || option.action === 'mergeAcross'
                    ? 'menuitem'
                    : 'menuitemradio'
                }
                aria-checked={
                  option.action === 'unmergeCells' || option.action === 'mergeAcross'
                    ? undefined
                    : checked
                }
                onClick={() => {
                  onPick(option.action);
                  setOpen(false);
                }}
              >
                <Icon name="merge" />
                <span>{option.label}</span>
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}
