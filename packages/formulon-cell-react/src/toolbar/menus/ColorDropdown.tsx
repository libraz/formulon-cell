// Font / fill color split-button that opens the shared color palette flyout.
// The palette widget is imperatively mounted (vanilla DOM) because the same
// implementation backs all three framework wrappers; React just owns the
// open/close lifecycle and the swatch button.

import { ChevronDown12Regular } from '@fluentui/react-icons';
import { createColorPalette } from '@libraz/formulon-cell';
import { type ReactElement, useEffect, useRef } from 'react';

import { RIBBON_KEYSHORTCUTS } from '../model.js';
import { useMenuOpen } from './useMenuOpen.js';

export interface ColorDropdownProps {
  id: string;
  title: string;
  value: string;
  labels: {
    automatic: string;
    moreColors: string;
    standardColors: string;
    themeColors: string;
  };
  label: ReactElement;
  disabled: boolean;
  onChange: (value: string) => void;
}

export function ColorDropdown({
  id,
  title,
  value,
  labels,
  label,
  disabled,
  onChange,
}: ColorDropdownProps): ReactElement {
  const { open, setOpen, wrapRef } = useMenuOpen();
  const hostRef = useRef<HTMLDivElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  // Latest props for the imperatively-mounted palette, so the mount effect
  // can depend on `[open]` alone and never re-create the widget mid-use.
  const latest = useRef({ id, title, value, labels, onChange });
  latest.current = { id, title, value, labels, onChange };

  useEffect(() => {
    const host = hostRef.current;
    if (!open || !host) return;
    const props = latest.current;
    const palette = createColorPalette({
      themeLabel: props.labels.themeColors,
      standardLabel: props.labels.standardColors,
      moreColorsLabel: props.labels.moreColors,
      ariaLabel: props.title,
      value: props.value,
      automatic:
        props.id === 'fontColor' ? { label: props.labels.automatic, color: '#000000' } : null,
      onPick: (color) => {
        latest.current.onChange(color);
        setOpen(false);
      },
      onMoreColors: () => {
        setOpen(false);
        inputRef.current?.click();
      },
    });
    host.appendChild(palette.el);
    palette.focus();
    return () => {
      palette.el.remove();
    };
  }, [open]);

  return (
    <div
      key={id}
      ref={wrapRef}
      className={`demo__rb-color${open ? ' demo__rb-color--open' : ''}`}
      data-ribbon-command={id}
      title={title}
    >
      <button
        type="button"
        className="demo__rb-color__btn"
        aria-label={title}
        aria-keyshortcuts={RIBBON_KEYSHORTCUTS[id]}
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((next) => !next)}
      >
        <span className="demo__rb-color__icon">{label}</span>
        <span className="demo__rb-color__swatch" style={{ backgroundColor: value }} />
        <ChevronDown12Regular className="demo__rb-color__chev" aria-hidden="true" />
      </button>
      {open ? <div ref={hostRef} className="demo__color-flyout" /> : null}
      <input
        ref={inputRef}
        className="demo__color-flyout__native"
        type="color"
        value={value}
        aria-hidden="true"
        tabIndex={-1}
        onChange={(e) => {
          onChange(e.currentTarget.value);
          setOpen(false);
        }}
      />
    </div>
  );
}
