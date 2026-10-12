import { type KeyboardEvent, type ReactElement, useId, useState } from 'react';
import { type DemoSelectOption, resolveDemoSelectKey } from '../../demo-shared/index.js';

interface DemoSelectProps {
  labelId: string;
  value: string;
  options: readonly DemoSelectOption[];
  onChange: (value: string) => void;
}

export const DemoSelect = ({
  labelId,
  value,
  options,
  onChange,
}: DemoSelectProps): ReactElement => {
  const listId = useId();
  const [open, setOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(-1);
  const selectedIndex = options.findIndex((o) => o.value === value);
  const optionId = (index: number): string => `${listId}-opt-${index}`;

  const openAt = (index: number): void => {
    setActiveIndex(index);
    setOpen(true);
  };
  const commit = (index: number): void => {
    const option = options[index];
    setOpen(false);
    if (option && option.value !== value) onChange(option.value);
  };

  const onKeyDown = (event: KeyboardEvent<HTMLButtonElement>): void => {
    const action = resolveDemoSelectKey(
      event.key,
      open,
      activeIndex,
      selectedIndex,
      options.length,
    );
    if (!action) return;
    if (action.kind !== 'commit' || !action.keepDefault) event.preventDefault();
    if (action.kind === 'open') openAt(action.index);
    else if (action.kind === 'move') setActiveIndex(action.index);
    else if (action.kind === 'commit') commit(action.index);
    else setOpen(false);
  };

  return (
    <div className="demo__select">
      <button
        type="button"
        role="combobox"
        className="demo__select-trigger"
        aria-labelledby={labelId}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-controls={listId}
        aria-activedescendant={open && activeIndex >= 0 ? optionId(activeIndex) : undefined}
        data-value={value}
        onClick={() => (open ? setOpen(false) : openAt(Math.max(0, selectedIndex)))}
        onKeyDown={onKeyDown}
        onBlur={() => setOpen(false)}
      >
        {options[selectedIndex]?.label ?? value}
      </button>
      <div
        id={listId}
        role="listbox"
        aria-labelledby={labelId}
        className="demo__select-list"
        hidden={!open}
      >
        {options.map((option, index) => (
          <button
            key={option.value}
            id={optionId(index)}
            type="button"
            tabIndex={-1}
            role="option"
            aria-selected={option.value === value}
            className={`demo__select-option${index === activeIndex ? ' demo__select-option--active' : ''}`}
            onMouseDown={(event) => event.preventDefault()}
            onMouseEnter={() => setActiveIndex(index)}
            onClick={() => commit(index)}
          >
            {option.label}
          </button>
        ))}
      </div>
    </div>
  );
};
