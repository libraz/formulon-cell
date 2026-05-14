import type {
  CellBorderStyle,
  MarginPreset,
  PageOrientation,
  PaperSize,
} from '@libraz/formulon-cell';
import { onMounted, onUnmounted, ref } from 'vue';
import type { BorderPreset } from './model.js';

type DropdownName =
  | 'fontFamily'
  | 'fontSize'
  | 'borderPreset'
  | 'borderStyle'
  | 'margins'
  | 'orientation'
  | 'paperSize';

interface DropdownHandlers {
  onBorderPreset(value: BorderPreset): void;
  onFontFamily(value: string): void;
  onFontSize(value: string | number): void;
  onMarginPreset(value: MarginPreset): void;
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
  const toggleDropdown = (name: DropdownName): void => {
    openDropdown.value = openDropdown.value === name ? null : name;
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
      closeDropdown();
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
  return { borderStyle, closeDropdown, onDropdownPick, openDropdown, toggleDropdown };
}
