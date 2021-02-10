import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { escape } from '@microsoft/sp-lodash-subset';

import TenantGlobalNavBar from './components/TenantGlobalNavBar';
import { ITenantGlobalNavbarProps } from './components/ITenantGlobalNavbarProps';
import * as SPTermStore from './services/SPTermStoreService';
import pnp from "sp-pnp-js";
import styles from './AppCustomizer.module.scss';



import * as strings from 'GlobalNavigationApplicationCustomizerStrings';
const LOG_SOURCE: string = 'GlobalNavigationApplicationCustomizer';
const NAV_TERMS_KEY: string = 'global-navigation-terms';

// Extra Import Need to remove Once development done
import { Dialog } from '@microsoft/sp-dialog';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGlobalNavigationApplicationCustomizerProperties {
  // This is an example; replace with your own property
  TopMenuTermSet?: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GlobalNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalNavigationApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _topMenuItems: SPTermStore.ISPTermObject[];

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    console.log("TopMenuTermsetName :"+this.properties.TopMenuTermSet);
    // Configure caching
    pnp.setup({
      defaultCachingStore: "session",
      defaultCachingTimeoutSeconds: 900, //15min
      globalCacheDisable: false // true to disable caching in case of debugging/testing
    });

    // Retrieve the menu items from taxonomy
    let termStoreService: SPTermStore.SPTermStoreService = new SPTermStore.SPTermStoreService({
      spHttpClient: this.context.spHttpClient,
      siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
    });

    if (this.properties.TopMenuTermSet == undefined || this.properties.TopMenuTermSet ==null) {
      let cachedTerms = pnp.storage.session.get(NAV_TERMS_KEY);
      if (cachedTerms != null) {
        this._topMenuItems = cachedTerms;
      }
      else {
        this._topMenuItems = await termStoreService.getTermsFromTermSetAsync("TenantGlobalNavBar", this.context.pageContext.web.language);
        pnp.storage.session.put(NAV_TERMS_KEY, this._topMenuItems);
      }
    }


    // Call render method for generating the needed html elements
    this._renderPlaceHolders();

    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {

    console.log('Available placeholders: ',
      this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this._topMenuItems != null && this._topMenuItems.length > 0) {
        const element: React.ReactElement<ITenantGlobalNavbarProps> = React.createElement(
          TenantGlobalNavBar,
          {
            menuItems: this._topMenuItems,
          }
        );
        ReactDom.render(element, this._topPlaceholder.domElement);
      }
    }

  }
  private _onDispose(): void {
    console.log('[TenantGlobalNavBarApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
