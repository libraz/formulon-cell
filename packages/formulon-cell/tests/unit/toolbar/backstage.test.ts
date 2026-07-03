import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import {
  backstageCardItems,
  backstageNavItems,
  createBackstageFactories,
} from '../../../src/toolbar/ribbon/backstage.js';

const ribbonText = dictionaries.en.ribbon;
const backstageText = dictionaries.en.backstage;

describe('toolbar/ribbon/backstage', () => {
  it('models the Excel-style File navigation and card actions in one shared list', () => {
    expect(backstageNavItems(backstageText, ribbonText).map((item) => item.action)).toEqual([
      'info',
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'share',
      'export',
      'options',
      'close',
    ]);

    expect(backstageCardItems(backstageText, ribbonText).map((item) => item.action)).toEqual([
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'page-setup',
      'edit-links',
      'share',
      'export',
      'options',
    ]);
  });

  it('renders Backstage actions from the shared model', () => {
    const docState = document.createElement('span');
    docState.textContent = 'Saved';
    const { createBackstageView } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      docState,
    });

    const view = createBackstageView();
    const titleMark = view.querySelector<HTMLElement>('.demo__backstage-xl');
    const preview = view.querySelector<HTMLElement>('.demo__backstage-preview');
    const backButton = view.querySelector<HTMLElement>('[data-backstage-action="back"]');
    const commandIcons = Array.from(
      view.querySelectorAll<HTMLElement>('.demo__backstage-command-icon'),
    );

    expect(
      Array.from(view.querySelectorAll<HTMLElement>('.demo__backstage-navitem')).map(
        (item) => item.dataset.backstageAction,
      ),
    ).toEqual([
      'back',
      'info',
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'share',
      'export',
      'options',
      'close',
    ]);
    expect(
      Array.from(view.querySelectorAll<HTMLElement>('.demo__backstage-card')).map(
        (item) => item.dataset.backstageAction,
      ),
    ).toEqual([
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'page-setup',
      'edit-links',
      'share',
      'export',
      'options',
    ]);
    expect(view.textContent).toContain('Workbook Information');
    expect(view.textContent).toContain('Saved');
    expect(backButton?.textContent).toBe('');
    expect(backButton?.classList.contains('demo__backstage-navitem--back')).toBe(true);
    expect(backButton?.getAttribute('aria-label')).toBe('Back');
    expect(backButton?.title).toBe('Back');
    expect(titleMark?.textContent).toBe('');
    expect(preview?.textContent).toBe('');
    expect(commandIcons.map((icon) => icon.textContent)).toEqual(['', '', '']);
    expect(commandIcons.map((icon) => icon.getAttribute('aria-hidden'))).toEqual([
      'true',
      'true',
      'true',
    ]);
    expect(commandIcons.map((icon) => icon.className)).toEqual([
      'demo__backstage-command-icon demo__backstage-command-icon--protect',
      'demo__backstage-command-icon demo__backstage-command-icon--inspect',
      'demo__backstage-command-icon demo__backstage-command-icon--manage',
    ]);
  });

  it('projects disabled reasons for unavailable Backstage card and command controls', () => {
    const { createBackstageCard, createBackstageCommand } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      docState: null,
    });

    const card = createBackstageCard('Host only', 'Requires integration', 'host-only', true);
    expect(card.disabled).toBe(true);
    expect(card.dataset.disabledReason).toBe(backstageText.commandUnavailable);
    expect(card.getAttribute('aria-description')).toBe(backstageText.commandUnavailable);

    const command = createBackstageCommand('Host command', 'Requires integration', 'H');
    expect(command.disabled).toBe(true);
    expect(command.dataset.disabledReason).toBe(backstageText.commandUnavailable);
  });

  it('centralizes Backstage action button creation', () => {
    const { createBackstagePrintAction, createBackstageButton, createBackstageCard } =
      createBackstageFactories({
        backstageText,
        ribbonText,
        shellSavedText: 'Ready',
        docName: () => 'Book1',
        docState: null,
      });

    const nav = createBackstageButton('Info', 'info', true, 'Workbook info');
    expect(nav.type).toBe('button');
    expect(nav.className).toBe('demo__backstage-navitem demo__backstage-navitem--active');
    expect(nav.dataset.backstageAction).toBe('info');
    expect(nav.getAttribute('aria-label')).toBe('Workbook info');

    const card = createBackstageCard('Open', 'Open file', 'open');
    expect(card.type).toBe('button');
    expect(card.className).toBe('demo__backstage-card');
    expect(card.dataset.backstageAction).toBe('open');

    const primary = createBackstagePrintAction('Print', 'print', true);
    expect(primary.type).toBe('button');
    expect(primary.className).toBe('demo__print-action demo__print-action--primary');
    expect(primary.dataset.backstageAction).toBe('print');
    expect(primary.textContent).toBe('Print');

    const secondary = createBackstagePrintAction('Page setup', 'page-setup');
    expect(secondary.type).toBe('button');
    expect(secondary.className).toBe('demo__print-action');
    expect(secondary.dataset.backstageAction).toBe('page-setup');
    expect(secondary.textContent).toBe('Page setup');
  });

  it('keeps Backstage Print action button DOM centralized', () => {
    const source = readFileSync(join(process.cwd(), 'src/toolbar/ribbon/backstage.ts'), 'utf8');
    const css = readFileSync(join(process.cwd(), 'src/styles/toolbar/base/backstage.css'), 'utf8');

    expect(source).toContain("import { createRibbonButton } from './button.js'");
    expect(source).toContain('const createBackstageActionButton');
    expect(source).toContain("createBackstageActionButton('demo__backstage-navitem'");
    expect(source).toContain("createBackstageActionButton('demo__backstage-card'");
    expect(source).toContain("'demo__backstage-command'");
    expect(source).toContain("createBackstageActionButton('demo__print-action'");
    expect(source).toContain('const createBackstagePrintAction');
    expect(source).toContain("createBackstagePrintAction(backstageText.printNow, 'print', true)");
    expect(source).toContain("createBackstagePrintAction(backstageText.printToPdf, 'export')");
    expect(source).toContain("createBackstagePrintAction(backstageText.pageSetup, 'page-setup')");
    expect(source).not.toContain("const card = document.createElement('button')");
    expect(source).not.toContain("const command = document.createElement('button')");
    expect(source).not.toContain("const print = document.createElement('button')");
    expect(source).not.toContain("const pdf = document.createElement('button')");
    expect(source).not.toContain("const pageSetup = document.createElement('button')");
    expect(source).not.toContain("document.createElement('button')");
    expect(source).not.toContain("preview.textContent = 'X'");
    expect(source).not.toContain("mark.textContent = 'X'");
    expect(source).not.toContain('mark.textContent = icon');
    expect(source).not.toContain("createBackstageButton('←'");
    expect(css).toMatch(
      /\.demo__backstage-navitem--back::after\s*\{[\s\S]*?border: solid currentColor;[\s\S]*?transform: translate\(-70%, -50%\) rotate\(45deg\);/,
    );
    expect(css).toMatch(
      /\.demo__backstage-xl\s*\{[\s\S]*?background: var\(--demo-brand\);[\s\S]*?color: var\(--demo-title-fg\);/,
    );
    expect(css).toMatch(
      /\.demo__backstage-xl \.demo__rb-icon\s*\{[\s\S]*?width: 30px;[\s\S]*?height: 30px;/,
    );
    expect(css).toMatch(
      /\.demo__backstage-command-icon--protect::before\s*\{[\s\S]*?border: 2px solid currentColor;/,
    );
    expect(css).toMatch(
      /\.demo__backstage-command-icon--inspect::before\s*\{[\s\S]*?background: #a4262c;/,
    );
    expect(css).toMatch(
      /\.demo__backstage-command-icon--manage::before\s*\{[\s\S]*?linear-gradient\(currentColor 0 0\)/,
    );
  });

  it('renders a dedicated Backstage Print view with command handoff attributes', () => {
    const { createBackstageView } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      docState: null,
    });

    const view = createBackstageView('print');
    expect(view.querySelector('[data-backstage-print-preview]')).toBeTruthy();
    expect(view.textContent).toContain('Print');
    expect(view.textContent).toContain('Export to PDF');
    expect(view.textContent).toContain('Page 1');
    expect(
      Array.from(view.querySelectorAll<HTMLButtonElement>('.demo__print-action')).map(
        (button) => button.dataset.backstageAction,
      ),
    ).toEqual(['print', 'export', 'page-setup']);
    const activePrint = Array.from(
      view.querySelectorAll<HTMLElement>('.demo__backstage-navitem--active'),
    ).map((button) => button.dataset.backstageAction);
    expect(activePrint).toEqual(['print']);
  });

  it('uses a host-supplied printable document for the shared Backstage Print preview', () => {
    const { createBackstageView } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      printPreviewHtml: () => '<!doctype html><html><body><main>Preview Cell</main></body></html>',
      docState: null,
    });

    const view = createBackstageView('print');
    const frame = view.querySelector<HTMLIFrameElement>('.demo__print-frame');

    expect(frame).toBeTruthy();
    expect(frame?.getAttribute('sandbox')).toBe('');
    expect(frame?.srcdoc).toContain('Preview Cell');
    expect(view.querySelector('.demo__print-page')).toBeNull();
  });
});
