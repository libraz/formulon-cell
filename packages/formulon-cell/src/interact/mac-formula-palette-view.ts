import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { makeButton, makeIconButton } from './mac-formula-palette-buttons.js';
import type { MacPaletteStrings } from './mac-formula-palette-catalog.js';

export interface PickerRefs {
  search: HTMLInputElement;
  sections: HTMLElement;
  insert: HTMLButtonElement;
}

export interface ArgumentsRefs {
  fields: HTMLElement;
  preview: HTMLElement;
  done: HTMLButtonElement;
}

export interface PickerRow {
  name: string;
  displayName: string;
  description: string;
  unavailable: boolean;
}

export interface PickerViewSection {
  key: string;
  title: string;
  rows: PickerRow[];
}

export interface PickerViewState {
  searchQuery: string;
  guardMessage: string;
  /** Canonical name of the highlighted row, if any. */
  activeName: string | null;
  sections: PickerViewSection[];
  summary: { displayName: string; description: string; syntax: string } | null;
  insertDisabled: boolean;
}

export interface ArgumentFieldView {
  label: string;
  description?: string;
  value: string;
}

export interface ArgumentsViewState {
  title: string;
  fields: ArgumentFieldView[];
  canAddArgument: boolean;
  doneDisabled: boolean;
  guardMessage: string;
  description: string;
  syntax: string;
  helpUrl?: string;
}

export interface MacFormulaPaletteViewContext {
  labels(): MacPaletteStrings;
  pickerState(): PickerViewState;
  argumentsState(): ArgumentsViewState;
  close(): void;
  showAll(): void;
  search(query: string): void;
  pick(name: string): void;
  insert(): void;
  argumentFocus(index: number): void;
  argumentInput(index: number, value: string): void;
  addArgument(): void;
  done(): void;
}

/** Palette DOM builders. Appends into `root`; the caller clears it first. */
export function createMacFormulaPaletteView(root: HTMLElement, ctx: MacFormulaPaletteViewContext) {
  let sectionsEl: HTMLElement | null = null;

  const appendPaletteHeader = (): void => {
    const labels = ctx.labels();
    const header = document.createElement('header');
    header.className = 'fc-mac-formula-palette__header';
    const title = document.createElement('h2');
    title.textContent = labels.title;
    header.appendChild(title);
    header.appendChild(makeIconButton(labels.close, 'close', 'close', () => ctx.close()));
    root.appendChild(header);
  };

  const appendGuard = (content: HTMLElement, message: string, role?: string): void => {
    if (!message) return;
    const guard = document.createElement('p');
    guard.className = 'fc-mac-formula-palette__guard';
    if (role) guard.dataset.role = role;
    guard.textContent = message;
    content.appendChild(guard);
  };

  const appendSection = (
    sections: HTMLElement,
    section: PickerViewSection,
    activeName: string | null,
  ): void => {
    const labels = ctx.labels();
    const el = document.createElement('section');
    el.className = 'fc-mac-formula-palette__section';
    el.dataset.section = section.key;
    const heading = document.createElement('h3');
    heading.textContent = section.title;
    el.appendChild(heading);
    const list = document.createElement('div');
    list.className = 'fc-mac-formula-palette__function-list';
    list.setAttribute('role', 'listbox');
    for (const entry of section.rows) {
      const row = document.createElement('div');
      row.className = 'fc-mac-formula-palette__function';
      row.dataset.functionName = entry.name;
      row.setAttribute('role', 'option');
      row.tabIndex = 0;
      row.setAttribute('aria-selected', String(activeName === entry.name));
      if (entry.unavailable) {
        row.classList.add('is-unavailable');
      }
      projectDisabledState(row, entry.unavailable, entry.unavailable ? labels.unavailable : null);
      const nameSpan = document.createElement('span');
      nameSpan.className = 'fc-mac-formula-palette__function-name';
      nameSpan.textContent = entry.displayName;
      row.appendChild(nameSpan);
      if (entry.description) {
        const descriptionSpan = document.createElement('span');
        descriptionSpan.className = 'fc-mac-formula-palette__function-description';
        descriptionSpan.textContent = entry.description;
        row.appendChild(descriptionSpan);
      }
      row.addEventListener('click', () => {
        if (!row.isConnected) return;
        ctx.pick(entry.name);
      });
      row.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          row.click();
        }
      });
      list.appendChild(row);
    }
    if (!list.childElementCount) {
      const empty = document.createElement('p');
      empty.className = 'fc-mac-formula-palette__empty';
      empty.textContent = labels.empty;
      list.appendChild(empty);
    }
    el.appendChild(list);
    sections.appendChild(el);
  };

  const renderPickerList = (): void => {
    const sections = sectionsEl;
    if (!sections) return;
    sections.replaceChildren();
    const state = ctx.pickerState();
    for (const section of state.sections) appendSection(sections, section, state.activeName);
  };

  const renderPickerSummary = (
    summary: HTMLElement,
    state: NonNullable<PickerViewState['summary']>,
  ): void => {
    summary.replaceChildren();
    const name = document.createElement('strong');
    name.textContent = state.displayName;
    summary.appendChild(name);
    if (state.description) {
      const text = document.createElement('span');
      text.textContent = state.description;
      summary.appendChild(text);
    }
    const syntax = document.createElement('code');
    syntax.textContent = state.syntax;
    summary.appendChild(syntax);
  };

  const renderPicker = (): PickerRefs => {
    const labels = ctx.labels();
    const state = ctx.pickerState();
    appendPaletteHeader();
    const content = document.createElement('div');
    content.className = 'fc-mac-formula-palette__content';
    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'fc-mac-formula-palette__search';
    search.placeholder = labels.searchPlaceholder;
    search.setAttribute('aria-label', labels.searchPlaceholder);
    search.value = state.searchQuery;
    search.addEventListener('input', () => ctx.search(search.value));
    content.appendChild(search);
    const sections = document.createElement('div');
    sections.className = 'fc-mac-formula-palette__sections';
    content.appendChild(sections);
    const summary = document.createElement('div');
    summary.className = 'fc-mac-formula-palette__summary';
    content.appendChild(summary);
    const insert = makeButton(labels.insertFunction, 'insert-function', () => ctx.insert());
    insert.className = 'fc-mac-formula-palette__insert';
    projectDisabledState(
      insert,
      state.insertDisabled,
      state.insertDisabled ? labels.unavailable : null,
    );
    content.appendChild(insert);
    appendGuard(content, state.guardMessage);
    root.appendChild(content);
    sectionsEl = sections;
    if (state.summary) renderPickerSummary(summary, state.summary);
    renderPickerList();
    return { search, sections, insert };
  };

  const renderArguments = (): ArgumentsRefs => {
    const labels = ctx.labels();
    const state = ctx.argumentsState();
    appendPaletteHeader();
    const content = document.createElement('div');
    content.className = 'fc-mac-formula-palette__content';
    const back = makeButton(labels.showAll, 'show-all', () => ctx.showAll());
    back.className = 'fc-mac-formula-palette__back';
    content.appendChild(back);

    const name = document.createElement('h3');
    name.className = 'fc-mac-formula-palette__args-name';
    name.textContent = state.title;
    content.appendChild(name);
    const fields = document.createElement('div');
    fields.className = 'fc-mac-formula-palette__fields';
    state.fields.forEach((help, index) => {
      const row = document.createElement('label');
      row.className = 'fc-mac-formula-palette__argument';
      const label = document.createElement('span');
      label.textContent = help.label;
      row.appendChild(label);
      const input = document.createElement('input');
      input.type = 'text';
      input.dataset.argumentIndex = String(index);
      input.value = help.value;
      input.addEventListener('focus', () => ctx.argumentFocus(index));
      input.addEventListener('input', () => ctx.argumentInput(index, input.value));
      row.appendChild(input);
      const range = makeIconButton(labels.rangePicker, 'range-picker', 'range', () => {
        ctx.argumentFocus(index);
        input.focus();
      });
      range.className = 'fc-range-picker__btn fc-mac-formula-palette__range';
      row.appendChild(range);
      if (help.description) {
        const hint = document.createElement('small');
        hint.textContent = help.description;
        row.appendChild(hint);
      }
      fields.appendChild(row);
    });
    content.appendChild(fields);
    if (state.canAddArgument) {
      const add = makeButton(labels.argument, 'add-argument', () => ctx.addArgument());
      add.className = 'fc-mac-formula-palette__add-argument';
      content.appendChild(add);
    }

    const resultRow = document.createElement('div');
    resultRow.className = 'fc-mac-formula-palette__result-row';
    const resultLabel = document.createElement('h4');
    resultLabel.textContent = labels.result;
    resultRow.appendChild(resultLabel);
    const preview = document.createElement('output');
    preview.className = 'fc-mac-formula-palette__preview';
    preview.dataset.role = 'preview-value';
    resultRow.appendChild(preview);
    const done = makeButton(labels.done, 'done', () => ctx.done());
    done.className = 'fc-mac-formula-palette__done';
    projectDisabledState(done, state.doneDisabled, state.doneDisabled ? labels.unavailable : null);
    resultRow.appendChild(done);
    content.appendChild(resultRow);
    appendGuard(content, state.guardMessage, 'draft-conflict');

    const help = document.createElement('section');
    help.className = 'fc-mac-formula-palette__help';
    const helpHeading = document.createElement('h4');
    helpHeading.textContent = labels.help;
    help.appendChild(helpHeading);
    const helpText = document.createElement('p');
    helpText.textContent = `${labels.description}: ${state.description}`;
    help.appendChild(helpText);
    const syntaxText = document.createElement('p');
    syntaxText.textContent = `${labels.syntax}: ${state.syntax}`;
    help.appendChild(syntaxText);
    if (state.helpUrl) {
      const link = document.createElement('a');
      link.href = state.helpUrl;
      link.target = '_blank';
      link.rel = 'noreferrer';
      link.textContent = labels.help;
      help.appendChild(link);
    }
    content.appendChild(help);
    root.appendChild(content);
    return { fields, preview, done };
  };

  return { renderPicker, renderPickerList, renderArguments };
}
