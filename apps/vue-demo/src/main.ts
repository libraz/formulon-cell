import '@libraz/formulon-cell/styles.css';
import '@libraz/formulon-cell/styles/contrast.css';
import './styles.css';

import { createApp } from 'vue';
import App from './App.vue';

const host = document.getElementById('app');
if (!host) throw new Error('#app missing');

createApp(App).mount(host);
