import '@libraz/formulon-cell/styles.css';
import '@libraz/formulon-cell/styles/contrast.css';
import './styles.css';
import '../../demo-shared/shell.css';

import { createApp } from 'vue';
import { installE2eProbes } from '../../demo-shared/e2e-probes.js';
import App from './App.vue';

const host = document.getElementById('app');
if (!host) throw new Error('#app missing');

installE2eProbes();
createApp(App).mount(host);
