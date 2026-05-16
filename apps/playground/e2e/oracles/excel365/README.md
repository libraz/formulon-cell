# Excel 365 UI Oracle

This folder keeps source-reference artifacts used while aligning the playground,
React, and Vue toolbar/dialog chrome toward Excel 365 behavior.

## Sources

- Ribbon display behavior and Ctrl+F1:
  https://support.microsoft.com/en-au/office/show-the-ribbon-26abd81c-b5ab-47a5-aabc-a9e5255862f4
- Backstage/File tab behavior:
  https://support.microsoft.com/office/start-backstage-with-the-file-tab-04610088-406c-43d0-98a0-c1999ab4ef53
- Mini toolbar behavior:
  https://support.microsoft.com/en-au/office/use-the-mini-toolbar-to-format-text-47012f83-b553-40a9-b7de-f038876f4db3
- Format Cells tabs and controls:
  https://learn.microsoft.com/en-US/office/troubleshoot/excel/format-cells-settings
- Format Cells visual reference:
  https://support.microsoft.com/en-us/office/format-numbers-f27f865b-2dc5-4970-b289-5286be8b994a
- Page Setup dialog tabs:
  https://support.microsoft.com/en-gb/office/page-setup-71c20d94-b13e-48fd-9800-cedd1fec6da3
- Fill handle and Auto Fill Options smart button:
  https://support.microsoft.com/en-gb/office/enter-a-series-of-numbers-dates-or-other-items-41e0bbf2-7198-4d78-8545-fdd4709976b4
- Paste Options and default paste behavior:
  https://support.microsoft.com/en-us/office/paste-options-8ea795b0-87cd-46af-9b59-ed4d8b1669ad
- Quick Analysis button and galleries:
  https://support.microsoft.com/en-us/office/basic-tasks-in-excel-dc775dd1-fa52-430f-9c3c-d998d1735fca
- Status bar options, default aggregates, and zoom controls:
  https://support.microsoft.com/en-us/office/excel-status-bar-options-6055ecd9-e20f-4a7a-a611-4481bd488c55
- Name Box cell/range selection and defined-name dropdown:
  https://support.microsoft.com/en-us/office/select-specific-cells-or-ranges-in-excel-3a0c91c5-8a64-4cd2-8625-7f5b7f1eed87
- Name Box naming a selected cell/range:
  https://support.microsoft.com/en-gb/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64
- Name Manager dialog, columns, and New/Edit/Delete/Filter behavior:
  https://support.microsoft.com/en-us/office/use-the-name-manager-in-excel-4d8c4c2b-9f7d-44e3-a3b4-9f61bd5c64e4
- Find and Replace dialog tabs, Options controls, Find All results, and Replace behavior:
  https://support.microsoft.com/en-us/office/find-or-replace-text-and-numbers-on-a-worksheet-0e304ca5-ecef-4808-b90f-fdb42f892e90
- Sheet tab rename and tab color behavior:
  https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
  https://support.microsoft.com/en-us/office/add-a-background-color-to-a-sheet-tab-440b28f2-3146-4dca-95df-3b9d43acbe59

## Saved Screenshots

- `ribbon-full-tabs-and-commands.png`:
  Microsoft Support image showing Excel ribbon tabs and commands.
- `backstage-file-options.png`:
  Microsoft Support image showing the File tab Backstage command list.
- `mini-toolbar.gif`:
  Microsoft Support image showing the Mini toolbar.
- `format-cells-number-tab.jpg`:
  Microsoft Support image showing the Format Cells dialog on the Number tab.
- `page-setup-page-options.png`, `page-setup-margin-options.png`,
  `page-setup-header-footer-options.png`, `page-setup-sheet-options.jpg`:
  Microsoft Support images showing the Page Setup dialog tabs.
- `auto-fill-options-button.gif`:
  Microsoft Support image showing the Auto Fill Options smart button.
- `paste-option-paste.png`, `paste-option-formatting.png`:
  Microsoft Support images for Paste Options commands.
- `quick-analysis-button.jpg`, `quick-analysis-selection.jpg`,
  `quick-analysis-totals-gallery.jpg`, `quick-analysis-formatting-gallery.jpg`:
  Microsoft Support images showing the Quick Analysis button and galleries.
- `sheet-tab-colors.png`:
  Microsoft Support image showing worksheet tabs with different tab colors.
- `name-box.png`:
  Microsoft Support image showing the Name Box to the left of the formula bar.
- `name-manager-dialog.png`:
  Microsoft Support image showing the Name Manager dialog box.
- `find-replace-options.png`, `find-replace-find-all.png`:
  Microsoft Support images showing the Excel Find and Replace dialog with Options
  and Find All results.

## Local Visual Coverage

- `../ribbon.spec.ts-snapshots/ribbon-home-chromium-darwin.png`
- `../ribbon.spec.ts-snapshots/ribbon-file-chromium-darwin.png`
- `../ribbon.spec.ts-snapshots/ribbon-collapsed-tabs-only-chromium-darwin.png`
- `../ribbon.spec.ts-snapshots/ribbon-display-options-menu-chromium-darwin.png`
- `../overlays.spec.ts-snapshots/overlays-context-menu-mini-toolbar-chromium-darwin.png`
- `../overlays.spec.ts-snapshots/overlays-auto-fill-options-date-menu-chromium-darwin.png`
- `../overlays.spec.ts-snapshots/overlays-paste-options-menu-chromium-darwin.png`
- `../overlays.spec.ts-snapshots/overlays-quick-analysis-panel-chromium-darwin.png`
- `../sheet-tabs.spec.ts-snapshots/sheet-tabs-colored-tab-chromium-darwin.png`
- `../sheet-tabs.spec.ts-snapshots/sheet-tabs-tab-color-menu-chromium-darwin.png`
- `../name-box.spec.ts-snapshots/name-box-defined-name-dropdown-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-format-cells-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-format-cells-number-category-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-format-cells-alignment-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-format-cells-font-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-page-setup-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-page-setup-margins-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-page-setup-header-footer-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-page-setup-sheet-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-find-replace-chromium-darwin.png`
- `../dialogs.spec.ts-snapshots/dialog-find-replace-options-chromium-darwin.png`
