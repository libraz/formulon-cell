// Backstage view ("File" tab) UI helpers extracted from main.ts. The view is
// purely presentational: it builds the navigation rail, properties panel, and
// the command/card grid. Behaviour is wired by the parent through data-backstage-*
// attributes on the produced buttons.
import { projectDisabledState } from '../menu-a11y.js';
import { createRibbonButton } from './button.js';

export interface BackstageText {
  properties: string;
  name: string;
  type: string;
  typeValue: string;
  status: string;
  location: string;
  locationValue: string;
  back: string;
  info: string;
  newLabel: string;
  open: string;
  save: string;
  saveAs: string;
  share: string;
  exportLabel: string;
  pageSetup: string;
  editLinks: string;
  options: string;
  close: string;
  subtitle: string;
  workbookInfo: string;
  protect: string;
  protectBody: string;
  inspect: string;
  inspectBody: string;
  manage: string;
  manageBody: string;
  newBody: string;
  openBody: string;
  saveBody: string;
  saveAsBody: string;
  printBody: string;
  printPreviewTitle: string;
  printNow: string;
  printToPdf: string;
  printSettings: string;
  printPreviewPage: string;
  printPreviewHint: string;
  shareBody: string;
  exportBody: string;
  pageSetupBody: string;
  editLinksBody: string;
  optionsBody: string;
  commandUnavailable: string;
}

export interface BackstageRibbonText {
  print: string;
  tabs: { file: string };
}

export interface BackstageDeps {
  backstageText: BackstageText;
  ribbonText: BackstageRibbonText;
  shellSavedText: string;
  /** The current document name (`docName` in main.ts). Captured at call-time. */
  docName: () => string;
  /** Optional host-supplied printable document HTML for the Backstage Print preview. */
  printPreviewHtml?: () => string | undefined;
  /** The doc-state DOM element used to derive a fallback status string. */
  docState: HTMLElement | null;
}

export interface BackstageFactories {
  createBackstageButton: (
    label: string,
    action: string,
    active?: boolean,
    ariaLabel?: string,
  ) => HTMLButtonElement;
  createBackstageCard: (
    title: string,
    body: string,
    action: string,
    disabled?: boolean,
    disabledReason?: string,
  ) => HTMLButtonElement;
  createBackstageCommand: (
    title: string,
    body: string,
    icon: string,
    action?: string,
  ) => HTMLButtonElement;
  createBackstagePrintAction: (
    label: string,
    action: BackstageAction,
    primary?: boolean,
  ) => HTMLButtonElement;
  createBackstageProperties: () => HTMLElement;
  createBackstagePrintView: () => HTMLElement;
  createBackstageView: (active?: BackstageAction) => HTMLElement;
}

export type BackstageAction =
  | 'back'
  | 'info'
  | 'new'
  | 'open'
  | 'save'
  | 'save-as'
  | 'print'
  | 'share'
  | 'export'
  | 'options'
  | 'close'
  | 'page-setup'
  | 'edit-links'
  | 'protect-workbook'
  | 'inspect-workbook';

export interface BackstageItem {
  label: string;
  action: BackstageAction;
  body?: string;
  active?: boolean;
}

export const backstageNavItems = (
  backstageText: BackstageText,
  ribbonText: BackstageRibbonText,
  active: BackstageAction = 'info',
): readonly BackstageItem[] => [
  { label: backstageText.info, action: 'info', active: active === 'info' },
  { label: backstageText.newLabel, action: 'new' },
  { label: backstageText.open, action: 'open' },
  { label: backstageText.save, action: 'save' },
  { label: backstageText.saveAs, action: 'save-as' },
  { label: ribbonText.print, action: 'print', active: active === 'print' },
  { label: backstageText.share, action: 'share' },
  { label: backstageText.exportLabel, action: 'export' },
  { label: backstageText.options, action: 'options' },
  { label: backstageText.close, action: 'close' },
];

export const backstageCardItems = (
  backstageText: BackstageText,
  ribbonText: BackstageRibbonText,
): readonly BackstageItem[] => [
  { label: backstageText.newLabel, action: 'new', body: backstageText.newBody },
  { label: backstageText.open, action: 'open', body: backstageText.openBody },
  { label: backstageText.save, action: 'save', body: backstageText.saveBody },
  { label: backstageText.saveAs, action: 'save-as', body: backstageText.saveAsBody },
  { label: ribbonText.print, action: 'print', body: backstageText.printBody },
  { label: backstageText.pageSetup, action: 'page-setup', body: backstageText.pageSetupBody },
  { label: backstageText.editLinks, action: 'edit-links', body: backstageText.editLinksBody },
  { label: backstageText.share, action: 'share', body: backstageText.shareBody },
  { label: backstageText.exportLabel, action: 'export', body: backstageText.exportBody },
  { label: backstageText.options, action: 'options', body: backstageText.optionsBody },
];

export const createBackstageFactories = (deps: BackstageDeps): BackstageFactories => {
  const { backstageText, ribbonText, shellSavedText, docName, printPreviewHtml, docState } = deps;

  const createBackstageActionButton = (
    className: string,
    action: string,
    label?: string,
    opts: { ariaLabel?: string; activeClassName?: string; active?: boolean } = {},
  ): HTMLButtonElement => {
    return createRibbonButton({
      className:
        opts.active && opts.activeClassName ? `${className} ${opts.activeClassName}` : className,
      dataset: { backstageAction: action },
      ariaLabel: opts.ariaLabel,
      text: label,
    });
  };

  const createBackstageButton = (
    label: string,
    action: string,
    active = false,
    ariaLabel?: string,
  ): HTMLButtonElement => {
    return createBackstageActionButton('fc-tb__backstage-navitem', action, label, {
      active,
      activeClassName: 'fc-tb__backstage-navitem--active',
      ariaLabel,
    });
  };

  const createBackstageCard = (
    title: string,
    body: string,
    action: string,
    disabled = false,
    disabledReason = backstageText.commandUnavailable,
  ): HTMLButtonElement => {
    const card = createBackstageActionButton('fc-tb__backstage-card', action);
    projectDisabledState(card, disabled, disabledReason, {
      datasetKey: 'disabledReason',
      titlePrefix: title,
    });
    const heading = document.createElement('strong');
    heading.textContent = title;
    const text = document.createElement('span');
    text.textContent = body;
    card.append(heading, text);
    return card;
  };

  const createBackstageCommand = (
    title: string,
    body: string,
    icon: string,
    action?: string,
  ): HTMLButtonElement => {
    const command = createBackstageActionButton(
      'fc-tb__backstage-command',
      action ?? 'unavailable',
    );
    if (!action) {
      delete command.dataset.backstageAction;
      projectDisabledState(command, true, backstageText.commandUnavailable, {
        datasetKey: 'disabledReason',
        titlePrefix: title,
      });
    }
    const mark = document.createElement('span');
    mark.className = `fc-tb__backstage-command-icon fc-tb__backstage-command-icon--${icon}`;
    mark.setAttribute('aria-hidden', 'true');
    const copy = document.createElement('span');
    const heading = document.createElement('strong');
    heading.textContent = title;
    const text = document.createElement('span');
    text.textContent = body;
    copy.append(heading, text);
    command.append(mark, copy);
    return command;
  };

  const createBackstagePrintAction = (
    label: string,
    action: BackstageAction,
    primary = false,
  ): HTMLButtonElement => {
    return createBackstageActionButton('fc-tb__print-action', action, label, {
      active: primary,
      activeClassName: 'fc-tb__print-action--primary',
    });
  };

  const createBackstageProperties = (): HTMLElement => {
    const props = document.createElement('aside');
    props.className = 'fc-tb__backstage-properties';
    const title = document.createElement('h2');
    title.className = 'fc-tb__backstage-section-title';
    title.textContent = backstageText.properties;
    const preview = document.createElement('div');
    preview.className = 'fc-tb__backstage-preview';
    const list = document.createElement('dl');
    list.className = 'fc-tb__backstage-prop-list';
    const pairs: [string, string][] = [
      [backstageText.name, docName()],
      [backstageText.type, backstageText.typeValue],
      [backstageText.status, docState?.textContent || shellSavedText],
      [backstageText.location, backstageText.locationValue],
    ];
    for (const [key, value] of pairs) {
      const dt = document.createElement('dt');
      dt.textContent = key;
      const dd = document.createElement('dd');
      dd.textContent = value;
      list.append(dt, dd);
    }
    props.append(title, preview, list);
    return props;
  };

  const createBackstagePrintView = (): HTMLElement => {
    const wrap = document.createElement('div');
    wrap.className = 'fc-tb__print-preview';
    wrap.dataset.backstagePrintPreview = 'true';

    const settings = document.createElement('section');
    settings.className = 'fc-tb__print-settings';
    settings.setAttribute('aria-label', backstageText.printSettings);
    const title = document.createElement('h2');
    title.textContent = backstageText.printPreviewTitle;
    const subtitle = document.createElement('p');
    subtitle.textContent = docName();
    const print = createBackstagePrintAction(backstageText.printNow, 'print', true);
    const pdf = createBackstagePrintAction(backstageText.printToPdf, 'export');
    const pageSetup = createBackstagePrintAction(backstageText.pageSetup, 'page-setup');
    settings.append(title, subtitle, print, pdf, pageSetup);

    const paper = document.createElement('section');
    paper.className = 'fc-tb__print-paper';
    paper.setAttribute('aria-label', backstageText.printPreviewPage);
    const previewHtml = printPreviewHtml?.();
    if (previewHtml) {
      const frame = document.createElement('iframe');
      frame.className = 'fc-tb__print-frame';
      frame.setAttribute('sandbox', '');
      frame.setAttribute('title', backstageText.printPreviewPage);
      frame.srcdoc = previewHtml;
      paper.appendChild(frame);
    } else {
      const page = document.createElement('div');
      page.className = 'fc-tb__print-page';
      const pageTitle = document.createElement('strong');
      pageTitle.textContent = `${backstageText.printPreviewPage} 1`;
      const lines = document.createElement('div');
      lines.className = 'fc-tb__print-sheet-lines';
      lines.setAttribute('aria-hidden', 'true');
      for (let i = 0; i < 12; i += 1) lines.appendChild(document.createElement('span'));
      page.append(pageTitle, lines);
      paper.appendChild(page);
    }
    const hint = document.createElement('p');
    hint.textContent = backstageText.printPreviewHint;
    paper.appendChild(hint);
    wrap.append(settings, paper);
    return wrap;
  };

  const createBackstageView = (active: BackstageAction = 'info'): HTMLElement => {
    const view = document.createElement('div');
    view.className = 'fc-tb__backstage';
    view.setAttribute('role', 'dialog');
    view.setAttribute('aria-modal', 'true');
    view.setAttribute('aria-label', ribbonText.tabs.file);

    const nav = document.createElement('nav');
    nav.className = 'fc-tb__backstage-nav';
    nav.setAttribute('aria-label', ribbonText.tabs.file);
    const title = document.createElement('strong');
    title.textContent = ribbonText.tabs.file;
    const backButton = createBackstageButton('', 'back', false, backstageText.back);
    backButton.classList.add('fc-tb__backstage-navitem--back');
    backButton.title = backstageText.back;
    nav.append(backButton, title);
    for (const item of backstageNavItems(backstageText, ribbonText, active)) {
      nav.appendChild(createBackstageButton(item.label, item.action, item.active));
    }

    const main = document.createElement('main');
    main.className = 'fc-tb__backstage-main';
    const heading = document.createElement('div');
    heading.className = 'fc-tb__backstage-title';
    const mark = document.createElement('span');
    mark.className = 'fc-tb__backstage-xl';
    mark.setAttribute('aria-hidden', 'true');
    const copy = document.createElement('div');
    const h1 = document.createElement('h1');
    h1.textContent = docName();
    const p = document.createElement('p');
    p.textContent = backstageText.subtitle;
    copy.append(h1, p);
    heading.append(mark, copy);

    const info = document.createElement('section');
    info.className = 'fc-tb__backstage-info';
    const manage = document.createElement('div');
    const manageTitle = document.createElement('h2');
    manageTitle.className = 'fc-tb__backstage-section-title';
    manageTitle.textContent = backstageText.workbookInfo;
    const commands = document.createElement('div');
    commands.className = 'fc-tb__backstage-command-list';
    commands.append(
      createBackstageCommand(
        backstageText.protect,
        backstageText.protectBody,
        'protect',
        'protect-workbook',
      ),
      createBackstageCommand(
        backstageText.inspect,
        backstageText.inspectBody,
        'inspect',
        'inspect-workbook',
      ),
      createBackstageCommand(backstageText.manage, backstageText.manageBody, 'manage', 'save-as'),
    );
    manage.append(manageTitle, commands);
    info.append(manage, createBackstageProperties());

    const grid = document.createElement('div');
    grid.className = 'fc-tb__backstage-grid';
    for (const item of backstageCardItems(backstageText, ribbonText)) {
      grid.appendChild(createBackstageCard(item.label, item.body ?? '', item.action));
    }

    if (active === 'print') main.append(heading, createBackstagePrintView());
    else main.append(heading, info, grid);
    view.append(nav, main);
    return view;
  };

  return {
    createBackstageButton,
    createBackstageCard,
    createBackstageCommand,
    createBackstagePrintAction,
    createBackstageProperties,
    createBackstagePrintView,
    createBackstageView,
  };
};
