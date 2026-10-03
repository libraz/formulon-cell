import {
  arrangeSessionIllustration,
  createRibbonImageFromSelection,
  createRibbonShapeFromSelection,
} from '../../commands/session-illustration.js';
import type { Range, SpreadsheetInstance } from '../../index.js';
import { pickImageFileDataUrl } from '../../toolbar/dialogs/image-file.js';
import type { DynamicDropdownsCtx } from '../../toolbar/ribbon/dynamic-dropdowns.js';

type IllustrationDropdownDefaultsDeps = {
  normalizedSelectionRange: (instance: SpreadsheetInstance) => Range;
  setMenuControlDisabled: (button: HTMLButtonElement, disabled: boolean, reason?: string) => void;
  showInstanceReport: (
    instance: SpreadsheetInstance,
    title: string,
    items: { severity: 'info' | 'warning'; label: string; detail: string }[],
  ) => Promise<void>;
};

type IllustrationDropdownDefaults = Pick<
  DynamicDropdownsCtx,
  | 'applyArrangeAction'
  | 'updateArrangeMenu'
  | 'insertPictureFromRibbon'
  | 'insertShapeFromRibbon'
  | 'insertScreenshotFromRibbon'
>;

export function createIllustrationDropdownDefaults(
  instance: SpreadsheetInstance,
  deps: IllustrationDropdownDefaultsDeps,
): IllustrationDropdownDefaults {
  const { normalizedSelectionRange, setMenuControlDisabled, showInstanceReport } = deps;

  const buildPictureAction =
    (instance: SpreadsheetInstance): DynamicDropdownsCtx['insertPictureFromRibbon'] =>
    async (action) => {
      const strings = instance.i18n.strings;
      if (action === 'device') {
        const result = await pickImageFileDataUrl();
        if (result) {
          createRibbonImageFromSelection(
            instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
            normalizedSelectionRange(instance),
            result.src,
            instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
            result.alt,
          );
        }
        instance.host.focus();
        return;
      }
      const ribbonMenu = strings.ribbonMenu as typeof strings.ribbonMenu & { pictureStock: string };
      const label =
        action === 'stock'
          ? ribbonMenu.pictureStock
          : action === 'online'
            ? strings.ribbonMenu.pictureOnline
            : strings.ribbonMenu.pictureThisDevice;
      const compatibilityDetails = strings.workbookObjects
        .compatibilityDetails as typeof strings.workbookObjects.compatibilityDetails & {
        mediaPickers?: string;
      };
      await showInstanceReport(instance, strings.ribbon.pictures, [
        {
          severity: 'info',
          label,
          detail: compatibilityDetails.mediaPickers ?? compatibilityDetails.chartsDrawings,
        },
      ]);
      instance.host.focus();
    };

  const buildShapeAction =
    (instance: SpreadsheetInstance): DynamicDropdownsCtx['insertShapeFromRibbon'] =>
    (shape) => {
      createRibbonShapeFromSelection(
        instance.store as unknown as Parameters<typeof createRibbonShapeFromSelection>[0],
        normalizedSelectionRange(instance),
        shape,
        instance.history as unknown as Parameters<typeof createRibbonShapeFromSelection>[3],
      );
      instance.host.focus();
    };

  const activeIllustrationId = (instance: SpreadsheetInstance): string | null => {
    const active = instance.host.querySelector<HTMLElement>(
      '.fc-illustration[aria-selected="true"]',
    );
    if (active?.dataset.illustrationId) return active.dataset.illustrationId;
    const state = instance.store.getState();
    const sheet = state.data.sheetIndex;
    const visible = state.illustrations.illustrations.filter((item) => item.sheet === sheet);
    return visible.at(-1)?.id ?? null;
  };

  const buildArrangeAction =
    (instance: SpreadsheetInstance): DynamicDropdownsCtx['applyArrangeAction'] =>
    (action) => {
      if (action === 'selection-pane') {
        instance.openWorkbookObjects();
        return;
      }
      const id = activeIllustrationId(instance);
      if (id) {
        arrangeSessionIllustration(
          instance.store as unknown as Parameters<typeof arrangeSessionIllustration>[0],
          id,
          action,
          instance.history as unknown as Parameters<typeof arrangeSessionIllustration>[3],
        );
      }
      instance.host.focus();
    };

  const updateArrangeMenu =
    (instance: SpreadsheetInstance): DynamicDropdownsCtx['updateArrangeMenu'] =>
    (menu) => {
      const activeId = activeIllustrationId(instance);
      const state = instance.store.getState();
      const sheet = state.data.sheetIndex;
      const visible = state.illustrations.illustrations.filter((item) => item.sheet === sheet);
      const activeIndex = activeId
        ? visible.findIndex((candidate) => candidate.id === activeId)
        : -1;
      const hasTarget = activeIndex >= 0;
      const atBack = !hasTarget || activeIndex === 0;
      const atFront = !hasTarget || activeIndex === visible.length - 1;
      const strings = instance.i18n.strings.ribbonMenu;
      for (const button of menu.querySelectorAll<HTMLButtonElement>('[data-arrange-action]')) {
        const action = button.dataset.arrangeAction;
        const disabled =
          (action === 'bring-forward' && atFront) ||
          (action === 'bring-front' && atFront) ||
          (action === 'send-backward' && atBack) ||
          (action === 'send-back' && atBack) ||
          (!hasTarget && action !== 'selection-pane');
        const reason = !hasTarget
          ? strings.arrangeRequiresObject
          : action === 'bring-forward' || action === 'bring-front'
            ? strings.arrangeAtFront
            : action === 'send-backward' || action === 'send-back'
              ? strings.arrangeAtBack
              : undefined;
        setMenuControlDisabled(button, disabled, reason);
      }
    };

  const buildScreenshotAction =
    (instance: SpreadsheetInstance): DynamicDropdownsCtx['insertScreenshotFromRibbon'] =>
    async (action = 'current-view') => {
      const strings = instance.i18n.strings;
      if (action === 'current-view') {
        const canvas = instance.host.querySelector<HTMLCanvasElement>('canvas');
        const dataUrl = canvas?.toDataURL?.('image/png');
        if (dataUrl) {
          createRibbonImageFromSelection(
            instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
            normalizedSelectionRange(instance),
            dataUrl,
            instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
          );
          instance.host.focus();
          return;
        }
      } else if (action === 'screen-clipping') {
        const captureScreenClip = (
          instance as unknown as {
            captureScreenClip: () => Promise<{ src: string; alt?: string } | null>;
          }
        ).captureScreenClip;
        const clip = await captureScreenClip();
        if (clip) {
          createRibbonImageFromSelection(
            instance.store as unknown as Parameters<typeof createRibbonImageFromSelection>[0],
            normalizedSelectionRange(instance),
            clip.src,
            instance.history as unknown as Parameters<typeof createRibbonImageFromSelection>[3],
            clip.alt,
          );
          instance.host.focus();
          return;
        }
      }
      const compatibilityDetails = strings.workbookObjects
        .compatibilityDetails as typeof strings.workbookObjects.compatibilityDetails & {
        screenshotCurrentView?: string;
        screenClipping?: string;
      };
      await showInstanceReport(instance, strings.ribbon.screenshot, [
        {
          severity: 'info',
          label:
            action === 'screen-clipping'
              ? strings.ribbonMenu.screenshotScreenClipping
              : strings.ribbonMenu.screenshotCurrentView,
          detail:
            action === 'screen-clipping'
              ? (compatibilityDetails.screenClipping ?? compatibilityDetails.chartsDrawings)
              : (compatibilityDetails.screenshotCurrentView ?? compatibilityDetails.chartsDrawings),
        },
      ]);
      instance.host.focus();
    };

  return {
    updateArrangeMenu: updateArrangeMenu(instance),
    applyArrangeAction: buildArrangeAction(instance),
    insertPictureFromRibbon: buildPictureAction(instance),
    insertShapeFromRibbon: buildShapeAction(instance),
    insertScreenshotFromRibbon: buildScreenshotAction(instance),
  };
}
