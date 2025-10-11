import { ApplicationConfig, provideBrowserGlobalErrorListeners, provideZoneChangeDetection } from '@angular/core';
import { provideRouter } from '@angular/router';

import { routes } from './app.routes';

import {
  HOT_GLOBAL_CONFIG,
  HotGlobalConfig,
  NON_COMMERCIAL_LICENSE,
} from "@handsontable/angular-wrapper";

import { registerLanguageDictionary, ruRU } from 'handsontable/i18n';
import { provideHttpClient } from '@angular/common/http';

registerLanguageDictionary(ruRU);


const globalHotConfig: HotGlobalConfig = {
  license: NON_COMMERCIAL_LICENSE,
  layoutDirection: "ltr",
  language: ruRU.languageCode,
  themeName: "ht-theme-main",
};


export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideZoneChangeDetection({ eventCoalescing: true }),
    provideRouter(routes),
    provideHttpClient(),
    { provide: HOT_GLOBAL_CONFIG, useValue: globalHotConfig },
  ]
};
