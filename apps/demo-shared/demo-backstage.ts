import type { PrinterProfile, SpreadsheetInstance } from '@libraz/formulon-cell';
import {
  buildPrintDocument,
  getPageSetup,
  resolvePrinterProfileBounds,
} from '@libraz/formulon-cell';
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
  action === 'export';

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
  else if (action === 'save' || action === 'save-as') opts.saveWorkbook();
  else if (action === 'print') instance?.print('print');
  else if (action === 'export') instance?.print('pdf');
  else if (action === 'page-setup') instance?.openPageSetup();
  else if (action === 'edit-links') instance?.openExternalLinksDialog();
  else if (action === 'share') opts.showNotice(opts.ui.share, opts.ui.shareDesc);
  else if (action === 'options') opts.toggleOptions();
  else if (action === 'close') opts.closeBackstage();
};

const formatMarginSummary = (margins: {
  top: number;
  right: number;
  bottom: number;
  left: number;
}): string => `${margins.top}" / ${margins.right}" / ${margins.bottom}" / ${margins.left}"`;

const selectedDemoPrinterProfile = (): PrinterProfile | undefined =>
  DEMO_PRINTER_PROFILES.find((profile) => profile.id === DEMO_PRINTER_PROFILE_ID) ??
  DEMO_PRINTER_PROFILES[0];

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
  const state = instance.store.getState();
  const sheet = state.data.sheetIndex;
  const setup = getPageSetup(state, sheet);
  const scale =
    setup.fitWidth || setup.fitHeight
      ? `${setup.fitWidth || 1} x ${setup.fitHeight || 1}`
      : `${Math.round((setup.scale ?? 1) * 100)}%`;
  const printerProfile = selectedDemoPrinterProfile();
  const printerBounds = resolvePrinterProfileBounds(
    setup,
    DEMO_PRINTER_PROFILES,
    DEMO_PRINTER_PROFILE_ID,
  );
  const printerSettings = [
    printerProfile?.name ? { label: ui.printPreviewPrinter, value: printerProfile.name } : null,
    printerBounds
      ? { label: ui.printPreviewPrinterMargins, value: formatMarginSummary(printerBounds) }
      : null,
  ].filter((item): item is { label: string; value: string } => item !== null);
  const printDocument = buildPrintDocument(
    instance.workbook,
    instance.store,
    sheet,
    ui.printPreviewTitle,
    {
      printableBounds: printerBounds ?? null,
    },
  );
  return {
    title: ui.printPreviewTitle,
    subtitle: bookName,
    printLabel: ui.printNow,
    pdfLabel: ui.printToPdf,
    pageSetupLabel: ui.pageSetup,
    previewTitle: `${ui.printPreviewPage} 1`,
    previewHint: ui.printPreviewHint,
    previewHtml: printDocument.html,
    settings: [
      { label: ui.printPreviewSheet, value: String(sheet + 1) },
      {
        label: ui.printPreviewOrientation,
        value:
          setup.orientation === 'portrait'
            ? ui.printPreviewOrientPortrait
            : ui.printPreviewOrientLandscape,
      },
      { label: ui.printPreviewPaper, value: setup.paperSize },
      ...printerSettings,
      { label: ui.printPreviewMargins, value: formatMarginSummary(setup.margins) },
      { label: ui.printPreviewScale, value: scale },
      { label: ui.printPreviewArea, value: setup.printArea?.trim() || ui.printPreviewNoArea },
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
];
