import type { PrinterProfile, SpreadsheetInstance } from '@libraz/formulon-cell';
import { formatA1Cell } from '@libraz/formulon-cell';
import { openDemoDialog } from './demo-actions.js';
import type { DemoUiStrings } from './demo-strings.js';

export const DEMO_PRINTER_PROFILES: readonly PrinterProfile[] = [
  {
    id: 'demo-office-a4',
    name: 'Demo Office Printer - A4',
    paperSize: 'A4',
    orientation: 'portrait',
    printableBounds: { top: 0.16, right: 0.14, bottom: 0.18, left: 0.14 },
  },
  {
    id: 'demo-office-a4-landscape',
    name: 'Demo Office Printer - A4 Landscape',
    paperSize: 'A4',
    orientation: 'landscape',
    printableBounds: { top: 0.14, right: 0.16, bottom: 0.14, left: 0.16 },
  },
  {
    id: 'demo-label-letter',
    name: 'Demo Label Printer - Letter',
    paperSize: 'letter',
    orientation: 'portrait',
    printableBounds: { top: 0.32, right: 0.28, bottom: 0.34, left: 0.28 },
  },
];

export const DEMO_PRINTER_PROFILE_ID = 'demo-office-a4';

export const refreshDemoPrinterProfiles = async (): Promise<readonly PrinterProfile[]> =>
  DEMO_PRINTER_PROFILES.map((profile) => ({
    ...profile,
    printableBounds: { ...profile.printableBounds },
  }));

export type DemoBackstageAction =
  | 'info'
  | 'new'
  | 'open'
  | 'save'
  | 'save-as'
  | 'print'
  | 'page-setup'
  | 'edit-links'
  | 'share'
  | 'export'
  | 'options'
  | 'autocorrect-options'
  | 'close';

export interface DemoBackstageItem {
  action: DemoBackstageAction;
  label: string;
  desc?: string;
  active?: boolean;
}

export interface DemoPrintPreviewModel {
  title: string;
  subtitle: string;
  printLabel: string;
  pdfLabel: string;
  pageSetupLabel: string;
  previewTitle: string;
  previewHint: string;
  previewHtml: string;
  settings: readonly { label: string; value: string }[];
}

export const DEMO_PRINT_PREVIEW_LINES = [
  'row-1',
  'row-2',
  'row-3',
  'row-4',
  'row-5',
  'row-6',
  'row-7',
  'row-8',
  'row-9',
  'row-10',
  'row-11',
  'row-12',
] as const;

export const demoBackstageRequiresWorkbook = (action: DemoBackstageAction): boolean =>
  action === 'save' ||
  action === 'save-as' ||
  action === 'print' ||
  action === 'page-setup' ||
  action === 'edit-links' ||
  action === 'export' ||
  action === 'autocorrect-options';

export const isDemoBackstageActionDisabled = (
  action: DemoBackstageAction,
  instance: SpreadsheetInstance | null | undefined,
): boolean => !instance && demoBackstageRequiresWorkbook(action);

export interface RunDemoBackstageActionOptions {
  action: DemoBackstageAction;
  instance: SpreadsheetInstance | null | undefined;
  ui: Pick<DemoUiStrings, 'share' | 'shareDesc'>;
  newWorkbook: () => void | Promise<void>;
  openWorkbook: () => void;
  saveWorkbook: () => void;
  saveWorkbookAs: () => void;
  showNotice: (title: string, detail: string) => void;
  toggleOptions: () => void;
  closeBackstage: () => void;
}

export const runDemoBackstageAction = (opts: RunDemoBackstageActionOptions): void => {
  const { action, instance } = opts;
  if (action === 'info') return;
  if (action === 'new') {
    void opts.newWorkbook();
  } else if (action === 'open') opts.openWorkbook();
  else if (action === 'save') opts.saveWorkbook();
  else if (action === 'save-as') opts.saveWorkbookAs();
  else if (action === 'print') instance?.print('print');
  else if (action === 'export') instance?.print('pdf');
  else if (action === 'page-setup') openDemoDialog(instance, 'pageSetup');
  else if (action === 'edit-links') openDemoDialog(instance, 'externalLinks');
  else if (action === 'share') opts.showNotice(opts.ui.share, opts.ui.shareDesc);
  else if (action === 'options') {
    // The Mac profile keeps its settings in the Preferences window; elsewhere the demo panel.
    if (instance?.dialogs.canOpen('preferences')) openDemoDialog(instance, 'preferences');
    else opts.toggleOptions();
  } else if (action === 'autocorrect-options') openDemoDialog(instance, 'autoCorrect');
  else if (action === 'close') opts.closeBackstage();
};

const formatMarginSummary = (margins: {
  top: number;
  right: number;
  bottom: number;
  left: number;
}): string => `${margins.top}" / ${margins.right}" / ${margins.bottom}" / ${margins.left}"`;

/** `PageOrientation.Landscape` of the engine's page setup. */
const ORIENTATION_LANDSCAPE = 1;

type DemoPaperSize = NonNullable<PrinterProfile['paperSize']>;

/** The OOXML paper size codes the spreadsheet names, and the ribbon string that labels each. */
const PAPER_SIZE_BY_CODE: Readonly<Record<number, DemoPaperSize>> = {
  1: 'letter',
  3: 'tabloid',
  5: 'legal',
  8: 'A3',
  9: 'A4',
  11: 'A5',
};

const PAPER_LABEL_KEYS = {
  A3: 'paperA3',
  A4: 'paperA4',
  A5: 'paperA5',
  letter: 'paperLetter',
  legal: 'paperLegal',
  tabloid: 'paperTabloid',
} as const satisfies Record<DemoPaperSize, string>;

const printableBound = (value: number | undefined): number =>
  Number.isFinite(value) ? Math.max(0, value ?? 0) : 0;

/** The printer profile the page setup prints on: the one the host selected by id, else the first
 *  that fits the paper size and orientation. */
const demoPrinterProfileFor = (
  instance: SpreadsheetInstance,
  paperSize: DemoPaperSize | undefined,
  landscape: boolean,
): PrinterProfile | undefined => {
  const { printerProfiles = [], printerProfileId } = instance.options.host ?? {};
  const orientation = landscape ? 'landscape' : 'portrait';
  return (
    printerProfiles.find((profile) => profile.id === printerProfileId) ??
    printerProfiles.find(
      (profile) =>
        (!profile.paperSize || profile.paperSize === paperSize) &&
        (!profile.orientation || profile.orientation === orientation),
    )
  );
};

export const buildDemoPrintPreviewModel = (
  ui: DemoUiStrings,
  instance: SpreadsheetInstance | null | undefined,
  bookName: string,
): DemoPrintPreviewModel => {
  if (!instance) {
    return {
      title: ui.print,
      subtitle: ui.printPreviewUnavailable,
      printLabel: ui.printNow,
      pdfLabel: ui.printToPdf,
      pageSetupLabel: ui.pageSetup,
      previewTitle: ui.printPreviewPage,
      previewHint: ui.printPreviewUnavailable,
      previewHtml: '',
      settings: [],
    };
  }
  const sheetIndex = instance.selection.get().sheet;
  const sheet = instance.workbook.sheet(sheetIndex);
  const setup = sheet.getPageSetup();
  const margins = sheet.getPageMargins();
  const scale = setup.fitToPage
    ? `${setup.fitToWidth || 1} x ${setup.fitToHeight || 1}`
    : `${setup.scalePct}%`;
  const landscape = setup.orientation === ORIENTATION_LANDSCAPE;
  const paperSize = PAPER_SIZE_BY_CODE[setup.paperSize];
  const printerProfile = demoPrinterProfileFor(instance, paperSize, landscape);
  const printerBounds = printerProfile && {
    top: printableBound(printerProfile.printableBounds.top),
    right: printableBound(printerProfile.printableBounds.right),
    bottom: printableBound(printerProfile.printableBounds.bottom),
    left: printableBound(printerProfile.printableBounds.left),
  };
  const printArea = sheet
    .getPrintArea()
    ?.map((r) => `${formatA1Cell(r.firstRow, r.firstCol)}:${formatA1Cell(r.lastRow, r.lastCol)}`)
    .join(',');
  const printerSettings = [
    printerProfile?.name ? { label: ui.printPreviewPrinter, value: printerProfile.name } : null,
    printerBounds
      ? { label: ui.printPreviewPrinterMargins, value: formatMarginSummary(printerBounds) }
      : null,
  ].filter((item): item is { label: string; value: string } => item !== null);
  return {
    title: ui.printPreviewTitle,
    subtitle: bookName,
    printLabel: ui.printNow,
    pdfLabel: ui.printToPdf,
    pageSetupLabel: ui.pageSetup,
    previewTitle: `${ui.printPreviewPage} 1`,
    previewHint: ui.printPreviewHint,
    previewHtml: instance.printDocument().html,
    settings: [
      { label: ui.printPreviewSheet, value: String(sheetIndex + 1) },
      {
        label: ui.printPreviewOrientation,
        value: landscape ? ui.printPreviewOrientLandscape : ui.printPreviewOrientPortrait,
      },
      {
        label: ui.printPreviewPaper,
        value: paperSize
          ? instance.i18n.strings.ribbon[PAPER_LABEL_KEYS[paperSize]]
          : `#${setup.paperSize}`,
      },
      ...printerSettings,
      {
        label: ui.printPreviewMargins,
        value: formatMarginSummary({
          top: margins.topIn,
          right: margins.rightIn,
          bottom: margins.bottomIn,
          left: margins.leftIn,
        }),
      },
      { label: ui.printPreviewScale, value: scale },
      { label: ui.printPreviewArea, value: printArea || ui.printPreviewNoArea },
    ],
  };
};

export const buildDemoBackstageNav = (
  ui: DemoUiStrings,
  active: DemoBackstageAction = 'info',
): readonly DemoBackstageItem[] => [
  { action: 'info', label: ui.info, active: active === 'info' },
  { action: 'new', label: ui.newWorkbook },
  { action: 'open', label: ui.openTitle },
  { action: 'save', label: ui.save },
  { action: 'save-as', label: ui.saveCopy },
  { action: 'print', label: ui.print, active: active === 'print' },
  { action: 'share', label: ui.share },
  { action: 'export', label: ui.export },
  { action: 'options', label: ui.options },
  { action: 'close', label: ui.close },
];

export const buildDemoBackstageCards = (ui: DemoUiStrings): readonly DemoBackstageItem[] => [
  { action: 'new', label: ui.newWorkbook, desc: ui.newWorkbookDesc },
  { action: 'open', label: ui.openTitle, desc: ui.openDesc },
  { action: 'save', label: ui.save, desc: ui.saveDesc },
  { action: 'save-as', label: ui.saveCopy, desc: ui.saveAsDesc },
  { action: 'print', label: ui.print, desc: ui.printDesc },
  { action: 'page-setup', label: ui.pageSetup, desc: ui.pageSetupDesc },
  { action: 'edit-links', label: ui.editLinks, desc: ui.linksDesc },
  { action: 'share', label: ui.share, desc: ui.shareDesc },
  { action: 'export', label: ui.export, desc: ui.exportDesc },
  { action: 'options', label: ui.options, desc: ui.optionsDesc },
  { action: 'autocorrect-options', label: ui.autoCorrectOptions, desc: ui.autoCorrectOptionsDesc },
];
