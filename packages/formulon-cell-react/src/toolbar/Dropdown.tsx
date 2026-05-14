import { Checkmark16Regular, ChevronDown12Regular } from '@fluentui/react-icons';
import { type ReactElement, useEffect, useRef, useState } from 'react';

interface DropdownOption<V extends string | number> {
  value: V;
  label: string;
}

interface DropdownProps<V extends string | number> {
  title: string;
  value: V;
  options: readonly DropdownOption<V>[];
  onChange: (value: V) => void;
  disabled?: boolean;
  className?: string;
  /** Optional override of what's shown in the closed display. Defaults to the
   *  option label matching `value`, or the raw `value` if no label matches.
   *  Used for the font-name dropdown so unknown faces still render. */
  display?: string;
}

export function Dropdown<V extends string | number>({
  title,
  value,
  options,
  onChange,
  disabled,
  className,
  display,
}: DropdownProps<V>): ReactElement {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  const matched = options.find((o) => o.value === value);
  const shown = display ?? matched?.label ?? String(value);

  useEffect(() => {
    if (!open) return;
    const onDocDown = (e: MouseEvent): void => {
      const node = wrapRef.current;
      if (!node) return;
      if (e.target instanceof Node && node.contains(e.target)) return;
      setOpen(false);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        e.preventDefault();
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  // Scroll the active row into view when opening, so long lists (font sizes)
  // don't strand the user at the top.
  useEffect(() => {
    if (!open) return;
    const list = listRef.current;
    if (!list) return;
    const sel = list.querySelector<HTMLElement>('[aria-selected="true"]');
    sel?.scrollIntoView({ block: 'nearest' });
  }, [open]);

  return (
    <div
      ref={wrapRef}
      className={`demo__rb-dd${className ? ` ${className}` : ''}${
        open ? ' demo__rb-dd--open' : ''
      }`}
    >
      <button
        type="button"
        className="demo__rb-dd__btn"
        title={title}
        aria-label={title}
        aria-haspopup="listbox"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((o) => !o)}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            setOpen(true);
          }
        }}
      >
        <span className="demo__rb-dd__value">{shown}</span>
        <ChevronDown12Regular className="demo__rb-dd__chev" />
      </button>
      {open ? (
        <div
          ref={listRef}
          className="demo__rb-dd__list"
          role="listbox"
          aria-label={title}
          tabIndex={-1}
        >
          {options.map((o) => {
            const selected = o.value === value;
            return (
              <button
                key={o.value}
                type="button"
                role="option"
                aria-selected={selected}
                className={`demo__rb-dd__opt${selected ? ' demo__rb-dd__opt--selected' : ''}`}
                onClick={() => {
                  onChange(o.value);
                  setOpen(false);
                }}
              >
                <span className="demo__rb-dd__check" aria-hidden="true">
                  {selected ? <Checkmark16Regular /> : null}
                </span>
                <span className="demo__rb-dd__label">{o.label}</span>
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}
