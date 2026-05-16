// Backstage view ("File" tab) UI helpers extracted from main.ts. The view is
// purely presentational: it builds the navigation rail, properties panel, and
// the command/card grid. Behaviour is wired by the parent through data-backstage-*
// attributes on the produced buttons.

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
  options: string;
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
  optionsBody: string;
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
  ) => HTMLButtonElement;
  createBackstageCommand: (
    title: string,
    body: string,
    icon: string,
    action?: string,
  ) => HTMLButtonElement;
  createBackstageProperties: () => HTMLElement;
  createBackstageView: () => HTMLElement;
}

export const createBackstageFactories = (deps: BackstageDeps): BackstageFactories => {
  const { backstageText, ribbonText, shellSavedText, docName, docState } = deps;

  const createBackstageButton = (
    label: string,
    action: string,
    active = false,
    ariaLabel?: string,
  ): HTMLButtonElement => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `demo__backstage-navitem${active ? ' demo__backstage-navitem--active' : ''}`;
    button.dataset.backstageAction = action;
    if (ariaLabel) button.setAttribute('aria-label', ariaLabel);
    button.textContent = label;
    return button;
  };

  const createBackstageCard = (
    title: string,
    body: string,
    action: string,
    disabled = false,
  ): HTMLButtonElement => {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'demo__backstage-card';
    card.dataset.backstageAction = action;
    card.disabled = disabled;
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
    const command = document.createElement('button');
    command.type = 'button';
    command.className = 'demo__backstage-command';
    if (action) command.dataset.backstageAction = action;
    else command.disabled = true;
    const mark = document.createElement('span');
    mark.className = 'demo__backstage-command-icon';
    mark.textContent = icon;
    const copy = document.createElement('span');
    const heading = document.createElement('strong');
    heading.textContent = title;
    const text = document.createElement('span');
    text.textContent = body;
    copy.append(heading, text);
    command.append(mark, copy);
    return command;
  };

  const createBackstageProperties = (): HTMLElement => {
    const props = document.createElement('aside');
    props.className = 'demo__backstage-properties';
    const title = document.createElement('h2');
    title.className = 'demo__backstage-section-title';
    title.textContent = backstageText.properties;
    const preview = document.createElement('div');
    preview.className = 'demo__backstage-preview';
    preview.textContent = 'X';
    const list = document.createElement('dl');
    list.className = 'demo__backstage-prop-list';
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

  const createBackstageView = (): HTMLElement => {
    const view = document.createElement('div');
    view.className = 'demo__backstage app__backstage';
    view.setAttribute('role', 'dialog');
    view.setAttribute('aria-modal', 'true');
    view.setAttribute('aria-label', ribbonText.tabs.file);

    const nav = document.createElement('nav');
    nav.className = 'demo__backstage-nav';
    nav.setAttribute('aria-label', ribbonText.tabs.file);
    const title = document.createElement('strong');
    title.textContent = ribbonText.tabs.file;
    nav.append(
      createBackstageButton('←', 'back', false, backstageText.back),
      title,
      createBackstageButton(backstageText.info, 'info', true),
      createBackstageButton(backstageText.newLabel, 'new'),
      createBackstageButton(backstageText.open, 'open'),
      createBackstageButton(backstageText.save, 'save'),
      createBackstageButton(backstageText.saveAs, 'save-as'),
      createBackstageButton(ribbonText.print, 'print'),
      createBackstageButton(backstageText.options, 'options'),
    );

    const main = document.createElement('main');
    main.className = 'demo__backstage-main';
    const heading = document.createElement('div');
    heading.className = 'demo__backstage-title';
    const mark = document.createElement('span');
    mark.className = 'demo__backstage-xl';
    mark.textContent = 'X';
    const copy = document.createElement('div');
    const h1 = document.createElement('h1');
    h1.textContent = docName();
    const p = document.createElement('p');
    p.textContent = backstageText.subtitle;
    copy.append(h1, p);
    heading.append(mark, copy);

    const info = document.createElement('section');
    info.className = 'demo__backstage-info';
    const manage = document.createElement('div');
    const manageTitle = document.createElement('h2');
    manageTitle.className = 'demo__backstage-section-title';
    manageTitle.textContent = backstageText.workbookInfo;
    const commands = document.createElement('div');
    commands.className = 'demo__backstage-command-list';
    commands.append(
      createBackstageCommand(
        backstageText.protect,
        backstageText.protectBody,
        'P',
        'protect-workbook',
      ),
      createBackstageCommand(
        backstageText.inspect,
        backstageText.inspectBody,
        '!',
        'inspect-workbook',
      ),
      createBackstageCommand(backstageText.manage, backstageText.manageBody, 'S', 'save-as'),
    );
    manage.append(manageTitle, commands);
    info.append(manage, createBackstageProperties());

    const grid = document.createElement('div');
    grid.className = 'demo__backstage-grid';
    grid.append(
      createBackstageCard(backstageText.newLabel, backstageText.newBody, 'new'),
      createBackstageCard(backstageText.open, backstageText.openBody, 'open'),
      createBackstageCard(backstageText.save, backstageText.saveBody, 'save'),
      createBackstageCard(backstageText.saveAs, backstageText.saveAsBody, 'save-as'),
      createBackstageCard(ribbonText.print, backstageText.printBody, 'print'),
      createBackstageCard(backstageText.options, backstageText.optionsBody, 'options'),
    );

    main.append(heading, info, grid);
    view.append(nav, main);
    return view;
  };

  return {
    createBackstageButton,
    createBackstageCard,
    createBackstageCommand,
    createBackstageProperties,
    createBackstageView,
  };
};
