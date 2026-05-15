import type {
  CellBorderStyle,
  MarginPreset,
  PageOrientation,
  PaperSize,
} from '@libraz/formulon-cell';
import { nextTick, onMounted, onUnmounted, ref } from 'vue';
import type { BorderPreset } from './model.js';

type DropdownName =
  | 'fontFamily'
  | 'fontSize'
  | 'borderPreset'
  | 'borderStyle'
  | 'fillColor'
  | 'fontColor'
  | 'margins'
  | 'merge'
  | 'numberFormat'
  | 'orientation'
  | 'paperSize';

interface DropdownHandlers {
  onBorderPreset(value: BorderPreset): void;
  onFontFamily(value: string): void;
  onFontSize(value: string | number): void;
  onMarginPreset(value: MarginPreset): void;
  onNumberFormat(value: string): void;
  onOpenPageSetup(): void;
  onPageOrientation(value: PageOrientation): void;
  onPaperSize(value: PaperSize): void;
}

export function useToolbarDropdown(handlers: DropdownHandlers) {
  const borderStyle = ref<CellBorderStyle>('thin');
  const openDropdown = ref<DropdownName | null>(null);
  const closeDropdown = (): void => {
    openDropdown.value = null;
  };
  const focusDropdownButton = (root: HTMLElement | null): void => {
    root?.querySelector<HTMLButtonElement>('.demo__rb-dd__btn')?.focus({ preventScroll: true });
  };
  const focusDropdownOption = (list: HTMLElement | null, index: number): void => {
    if (!list) return;
    const options = Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]'));
    if (options.length === 0) return;
    const next = ((index % options.length) + options.length) % options.length;
    for (const [idx, option] of options.entries()) option.tabIndex = idx === next ? 0 : -1;
    options[next]?.focus({ preventScroll: true });
    options[next]?.scrollIntoView({ block: 'nearest' });
  };
  const focusOpenDropdownSelection = async (): Promise<void> => {
    await nextTick();
    const root = document.querySelector<HTMLElement>('.demo__rb-dd--open');
    const list = root?.querySelector<HTMLElement>('.demo__rb-dd__list') ?? null;
    const selectedIndex = Math.max(
      0,
      Array.from(list?.querySelectorAll<HTMLButtonElement>('[role="option"]') ?? []).findIndex(
        (option) => option.getAttribute('aria-selected') === 'true',
      ),
    );
    focusDropdownOption(list, selectedIndex);
  };
  const toggleDropdown = (name: DropdownName): void => {
    openDropdown.value = openDropdown.value === name ? null : name;
    if (openDropdown.value === name) void focusOpenDropdownSelection();
  };
  const onDropdownPick = (name: DropdownName, value: string | number): void => {
    if (name === 'fontFamily') handlers.onFontFamily(String(value));
    else if (name === 'fontSize') handlers.onFontSize(value);
    else if (name === 'borderPreset') handlers.onBorderPreset(String(value) as BorderPreset);
    else if (name === 'borderStyle') borderStyle.value = String(value) as CellBorderStyle;
    else if (name === 'margins') {
      if (value === 'custom') handlers.onOpenPageSetup();
      else handlers.onMarginPreset(String(value) as MarginPreset);
    } else if (name === 'orientation') handlers.onPageOrientation(String(value) as PageOrientation);
    else if (name === 'paperSize') handlers.onPaperSize(String(value) as PaperSize);
    else if (name === 'numberFormat') handlers.onNumberFormat(String(value));
    closeDropdown();
  };
  const onDocPointerDown = (e: MouseEvent): void => {
    if (openDropdown.value == null) return;
    const target = e.target;
    if (!(target instanceof Element)) return;
    if (target.closest('.demo__rb-dd')) return;
    closeDropdown();
  };
  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape' && openDropdown.value != null) {
      e.preventDefault();
      const root = document.querySelector<HTMLElement>('.demo__rb-dd--open');
      closeDropdown();
      focusDropdownButton(root);
    }
  };
  const onDropdownKeydown = (e: KeyboardEvent): void => {
    const target = e.target as Element | null;
    const root = target?.closest<HTMLElement>('.demo__rb-dd');
    if (!root) return;
    const name = root.dataset.dropdownName as DropdownName | undefined;
    if (!name) return;
    const button = target?.closest<HTMLButtonElement>('.demo__rb-dd__btn');
    if (button) {
      if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openDropdown.value = name;
        void focusOpenDropdownSelection();
      } else if (e.key === 'Escape' && openDropdown.value === name) {
        e.preventDefault();
        closeDropdown();
        button.focus({ preventScroll: true });
      }
      return;
    }

    const list = target?.closest<HTMLElement>('.demo__rb-dd__list');
    if (!list) return;
    const options = Array.from(list.querySelectorAll<HTMLButtonElement>('[role="option"]'));
    const current = Math.max(0, options.indexOf(document.activeElement as HTMLButtonElement));
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusDropdownOption(list, current + 1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusDropdownOption(list, current - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusDropdownOption(list, 0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusDropdownOption(list, options.length - 1);
    } else if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      (document.activeElement as HTMLButtonElement | null)?.click();
      focusDropdownButton(root);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      closeDropdown();
      focusDropdownButton(root);
    }
  };
  onMounted(() => {
    document.addEventListener('mousedown', onDocPointerDown, true);
    document.addEventListener('keydown', onDocKey, true);
  });
  onUnmounted(() => {
    document.removeEventListener('mousedown', onDocPointerDown, true);
    document.removeEventListener('keydown', onDocKey, true);
  });
  return {
    borderStyle,
    closeDropdown,
    onDropdownKeydown,
    onDropdownPick,
    openDropdown,
    toggleDropdown,
  };
}
