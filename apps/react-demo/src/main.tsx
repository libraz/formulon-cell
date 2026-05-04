import '@libraz/formulon-cell/styles.css';
import '@libraz/formulon-cell/styles/contrast.css';
import './styles.css';

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App.js';

const host = document.getElementById('root');
if (!host) throw new Error('#root missing');

createRoot(host).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
