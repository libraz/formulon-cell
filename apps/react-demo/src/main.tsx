import '@libraz/formulon-cell/styles.css';
import '@libraz/formulon-cell/styles/contrast.css';
import './styles.css';
import '../../demo-shared/shell.css';

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { installE2eProbes } from '../../demo-shared/e2e-probes.js';
import { App } from './App.js';

const host = document.getElementById('root');
if (!host) throw new Error('#root missing');

installE2eProbes();
createRoot(host).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
