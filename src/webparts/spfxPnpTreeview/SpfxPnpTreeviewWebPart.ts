import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxPnpTreeviewWebPartStrings';
import SpfxPnpTreeview from './components/SpfxPnpTreeview';
import { ISpfxPnpTreeviewProps } from './components/ISpfxPnpTreeviewProps';
import { IDynamicDataPropertyDefinition, IDynamicDataCallables } from '@microsoft/sp-dynamic-data';
import { sp } from '@pnp/sp';

export interface ISpfxPnpTreeviewWebPartProps {
  title: string;
}

export interface IWikiPage {
  wikiPageTitle: string;
}

export default class SpfxPnpTreeviewWebPart extends BaseClientSideWebPart<ISpfxPnpTreeviewWebPartProps> implements IDynamicDataCallables {

  private _selectedWikiPage: IWikiPage;

  private _wikiPageSelected = (wikiPage: IWikiPage): void => {
    this._selectedWikiPage = wikiPage;
    this.context.dynamicDataSourceManager.notifyPropertyChanged('wikiPage');
  }

  protected onInit(): Promise<void> {
    // setup PnPjs context
    sp.setup({
      spfxContext: this.context
    });
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve();
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'wikiPage',
        title: 'WikiPage'
      }
    ];
  }

  public getPropertyValue(propertyId: string): IWikiPage {
    switch (propertyId) {
      case 'wikiPage':
        return this._selectedWikiPage;
    }
    throw new Error('Bad property id');
  }

  public render(): void {
    const element: React.ReactElement<ISpfxPnpTreeviewProps> = React.createElement(
      SpfxPnpTreeview,
      {
        displayMode: this.displayMode,
        onWikiPageSelected: this._wikiPageSelected,
        title: this.properties.title,
        updateProperty: (value: string): void => {
          this.properties.title = value;
        },
        siteUrl: this.context.pageContext.web.serverRelativeUrl
        //context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 label: strings.DescriptionFieldLabel
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }
}
