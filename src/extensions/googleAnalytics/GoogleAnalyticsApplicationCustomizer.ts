import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'GoogleAnalyticsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AnalyticsApplicationCustomizer';

var currentURL: string = document.location.href;
var previousURL: string = "";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGoogleAnalyticsApplicationCustomizerProperties {
  trackingID: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GoogleAnalyticsApplicationCustomizer
  extends BaseApplicationCustomizer<IGoogleAnalyticsApplicationCustomizerProperties> {

  private currentPage = "";
  private isInitialLoad = true;

  private getFreshCurrentPage(): string {
    return window.location.pathname + window.location.search;
  }

  private updateCurrentPage(): void {
    this.currentPage = this.getFreshCurrentPage();
  }

  private navigatedEvent(): void {

    let trackingID: string = this.properties.trackingID;
    if (!trackingID) {
      Log.info(LOG_SOURCE, `${strings.MissingID}`);
    } else {

      const navigatedPage = this.getFreshCurrentPage();

      if (this.isInitialLoad) {
        this.realInitialNavigatedEvent(trackingID);
        this.updateCurrentPage();
        this.isInitialLoad = false;

      }
      else if (!this.isInitialLoad && (navigatedPage !== this.currentPage)) {
        this.realNavigatedEvent(trackingID);
        this.updateCurrentPage();
      }
    }
  }

  private realInitialNavigatedEvent(trackingID: string): void {
    console.log("Adding GTM full page load...");

    var gtmScript = document.createElement("script");
    gtmScript.type = "text/javascript";
    gtmScript.src = `https://www.googletagmanager.com/gtm.js?id=${trackingID}`;
    gtmScript.async = true;
    document.head.appendChild(gtmScript);

    eval(`
            window.dataLayer = window.dataLayer || [];
            window.dataLayer.push({'gtm.start':new Date().getTime(), event: 'gtm.js'});
        `);
  }

  private realNavigatedEvent(trackingID: string): void {

    console.log("Tracking partial page load...");

    eval(`
          window.dataLayer = window.dataLayer || [];
          window.dataLayer.push({'event':'partial-pageview', 'page': ${document.URL}});
        `);
  }

  @override
  public onInit(): Promise<void> {

    this.context.application.navigatedEvent.add(this, this.navigatedEvent);

    return Promise.resolve();
  }
}
