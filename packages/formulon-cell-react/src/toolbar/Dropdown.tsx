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
  const buttonRef = useRef<HTMLButtonElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  // Index of the option currently holding roving focus. Initialised to the
  // selected option each time the list opens.
  const [focusIdx, setFocusIdx] = useState(-1);
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
        buttonRef.current?.focus({ preventScroll: true });
      }
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  // On open: seed focus on the selected option and scroll it into view so
  // long lists (font sizes) don't strand the user at the top.
  useEffect(() => {
    if (!open) return;
    const idx = Math.max(
      0,
      options.findIndex((o) => o.value === value),
    );
    setFocusIdx(idx);
  }, [open, options, value]);

  // Move DOM focus + scroll-into-view whenever the focused index changes
  // while the list is open.
  useEffect(() => {
    if (!open || focusIdx < 0) return;
    const list = listRef.current;
    if (!list) return;
    const target = list.querySelectorAll<HTMLElement>('[role="option"]')[focusIdx];
    if (!target) return;
    target.focus();
    target.scrollIntoView({ block: 'nearest' });
  }, [open, focusIdx]);

  const moveFocus = (delta: number): void => {
    if (options.length === 0) return;
    setFocusIdx((cur) => {
      const start = cur < 0 ? 0 : cur;
      const next = start + delta;
      // Wrap around so ArrowUp at top jumps to bottom (Excel-like).
      if (next < 0) return options.length - 1;
      if (next >= options.length) return 0;
      return next;
    });
  };

  const onListKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      moveFocus(1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      moveFocus(-1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      setFocusIdx(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      setFocusIdx(options.length - 1);
    } else if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      const idx = focusIdx >= 0 ? focusIdx : options.findIndex((o) => o.value === value);
      const opt = options[idx];
      if (opt) {
        onChange(opt.value);
        setOpen(false);
        buttonRef.current?.focus({ preventScroll: true });
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setOpen(false);
      buttonRef.current?.focus({ preventScroll: true });
    }
  };

  return (
    <div
      ref={wrapRef}
      className={`demo__rb-dd${className ? ` ${className}` : ''}${
        open ? ' demo__rb-dd--open' : ''
      }`}
    >
      <button
        ref={buttonRef}
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
          onKeyDown={onListKeyDown}
        >
          {options.map((o, idx) => {
            const selected = o.value === value;
            return (
              <button
                key={o.value}
                type="button"
                role="option"
                aria-selected={selected}
                tabIndex={idx === focusIdx ? 0 : -1}
                className={`demo__rb-dd__opt${selected ? ' demo__rb-dd__opt--selected' : ''}`}
                onClick={() => {
                  onChange(o.value);
                  setOpen(false);
                  buttonRef.current?.focus({ preventScroll: true });
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
