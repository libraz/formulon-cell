import type { Page } from '@playwright/test';

/** Takes the demo page's ribbon and title bar out of rendering: the outermost ancestor of the
 *  ribbon tabs that does not hold the spreadsheet. */
export const hidePageRibbon = (page: Page): Promise<void> =>
  page.evaluate(() => {
    let chrome = document.querySelector<HTMLElement>('[data-ribbon-tab]');
    while (chrome?.parentElement && !chrome.parentElement.querySelector('.fc-host')) {
      chrome = chrome.parentElement;
    }
    if (chrome) chrome.style.display = 'none';
  });
