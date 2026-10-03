import { parseScriptCommand, type ScriptCommand } from '@libraz/formulon-cell';
import { type ReactElement, useRef, useState } from 'react';
import type { DemoCommandStrings, DemoUiStrings } from '../../demo-shared/index.js';
import { useDemoModalFocus } from './useDemoModalFocus.js';

interface DemoScriptDialogProps {
  ui: DemoUiStrings;
  commandText: DemoCommandStrings;
  /** Called with the parsed command; the host decides when to close. */
  onSubmit: (command: ScriptCommand) => void;
  /** Must be referentially stable: it keys the focus-trap effect. */
  onClose: () => void;
}

export const DemoScriptDialog = ({
  ui,
  commandText,
  onSubmit,
  onClose,
}: DemoScriptDialogProps): ReactElement => {
  const [command, setCommand] = useState('uppercase');
  const [error, setError] = useState<string | null>(null);
  const modalRef = useRef<HTMLDivElement | null>(null);
  useDemoModalFocus(modalRef, true, onClose);

  const submit = (): void => {
    const parsed = parseScriptCommand(command);
    if (!parsed) {
      setError(commandText.scriptCommandError);
      return;
    }
    onSubmit(parsed);
  };

  return (
    <div
      ref={modalRef}
      className="fc-tb__modal"
      role="dialog"
      aria-modal="true"
      aria-label={commandText.script}
    >
      <form
        className="fc-tb__modal-panel fc-tb__modal-panel--narrow"
        onSubmit={(ev) => {
          ev.preventDefault();
          submit();
        }}
      >
        <header className="fc-tb__modal-header">
          <h2>{commandText.script}</h2>
          <button type="button" className="fc-tb__modal-x" aria-label={ui.close} onClick={onClose}>
            ×
          </button>
        </header>
        <div className="fc-tb__modal-body">
          <label className="fc-tb__modal-field">
            <span>{ui.command}</span>
            <input
              value={command}
              onChange={(ev) => {
                setCommand(ev.target.value);
                setError(null);
              }}
            />
          </label>
          {error ? <p className="fc-tb__modal-error">{error}</p> : null}
        </div>
        <footer className="fc-tb__modal-footer">
          <button type="button" className="fc-tb__btn" onClick={onClose}>
            {ui.cancel}
          </button>
          <button type="submit" className="fc-tb__btn fc-tb__btn--active">
            {ui.run}
          </button>
        </footer>
      </form>
    </div>
  );
};
