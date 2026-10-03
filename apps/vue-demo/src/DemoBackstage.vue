<script setup lang="ts">
import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { computed } from 'vue';
import {
  buildDemoBackstageCards,
  buildDemoBackstageNav,
  buildDemoPrintPreviewModel,
  DEMO_PRINT_PREVIEW_LINES,
  type DemoBackstageAction,
  type DemoUiStrings,
  isDemoBackstageActionDisabled,
} from '../../demo-shared/index.js';
import DemoIcon from './DemoIcon.vue';

const props = defineProps<{
  ui: DemoUiStrings;
  instance: SpreadsheetInstance | null;
  bookName: string;
  action: DemoBackstageAction;
}>();
const emit = defineEmits<{ action: [action: DemoBackstageAction] }>();

const nav = computed(() => buildDemoBackstageNav(props.ui, props.action));
const cards = computed(() => buildDemoBackstageCards(props.ui));
const printPreview = computed(() => {
  void props.action;
  return buildDemoPrintPreviewModel(props.ui, props.instance, props.bookName);
});
</script>

<template>
  <div class="fc-tb__backstage" role="dialog" :aria-label="ui.file">
    <nav class="fc-tb__backstage-nav" :aria-label="ui.file">
      <strong>{{ ui.file }}</strong>
      <button
        v-for="item in nav"
        :key="item.action"
        type="button"
        :class="[
          'fc-tb__backstage-navitem',
          item.active ? 'fc-tb__backstage-navitem--active' : '',
        ]"
        :disabled="isDemoBackstageActionDisabled(item.action, instance)"
        @click="emit('action', item.action)"
      >
        {{ item.label }}
      </button>
    </nav>
    <div class="fc-tb__backstage-main">
      <div class="fc-tb__backstage-title">
        <span class="fc-tb__backstage-xl" aria-hidden="true">
          <DemoIcon name="app" />
        </span>
        <div>
          <h1>{{ bookName }}</h1>
          <p>{{ ui.backstageSub }}</p>
        </div>
      </div>
      <div v-if="action === 'print'" class="fc-tb__print-preview" data-demo-print-preview>
        <section class="fc-tb__print-settings" :aria-label="ui.printSettings">
          <h2>{{ printPreview.title }}</h2>
          <p>{{ printPreview.subtitle }}</p>
          <button
            type="button"
            class="fc-tb__print-action fc-tb__print-action--primary"
            :disabled="!instance"
            @click="instance?.print('print')"
          >
            {{ printPreview.printLabel }}
          </button>
          <button
            type="button"
            class="fc-tb__print-action"
            :disabled="!instance"
            @click="instance?.print('pdf')"
          >
            {{ printPreview.pdfLabel }}
          </button>
          <button
            type="button"
            class="fc-tb__print-action"
            :disabled="!instance"
            @click="instance?.openPageSetup()"
          >
            {{ printPreview.pageSetupLabel }}
          </button>
          <dl class="fc-tb__print-meta">
            <div v-for="row in printPreview.settings" :key="row.label">
              <dt>{{ row.label }}</dt>
              <dd>{{ row.value }}</dd>
            </div>
          </dl>
        </section>
        <section class="fc-tb__print-paper" :aria-label="printPreview.previewTitle">
          <iframe
            v-if="printPreview.previewHtml"
            class="fc-tb__print-frame"
            :title="printPreview.previewTitle"
            sandbox=""
            :srcdoc="printPreview.previewHtml"
          />
          <div v-else class="fc-tb__print-page">
            <strong>{{ printPreview.previewTitle }}</strong>
            <div aria-hidden="true" class="fc-tb__print-sheet-lines">
              <span v-for="line in DEMO_PRINT_PREVIEW_LINES" :key="line" />
            </div>
          </div>
          <p>{{ printPreview.previewHint }}</p>
        </section>
      </div>
      <div v-else class="fc-tb__backstage-grid">
        <button
          v-for="item in cards"
          :key="item.action"
          type="button"
          class="fc-tb__backstage-card"
          :disabled="isDemoBackstageActionDisabled(item.action, instance)"
          @click="emit('action', item.action)"
        >
          <strong>{{ item.label }}</strong>
          <span>{{ item.desc }}</span>
        </button>
      </div>
    </div>
  </div>
</template>
