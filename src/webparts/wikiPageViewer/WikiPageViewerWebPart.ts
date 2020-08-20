import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
// import {
//   IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-property-pane';
import { IPropertyPaneConfiguration, PropertyPaneDynamicFieldSet, PropertyPaneDynamicField,DynamicDataSharedDepth } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart,IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';

import * as strings from 'WikiPageViewerWebPartStrings';
import WikiPageViewer from './components/WikiPageViewer';
import { IWikiPageViewerProps } from './components/IWikiPageViewerProps';
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IWikiPageViewerWebPartProps {
 
  /**
   * wiki page to show the details for
   */
  wikiPage: DynamicProperty<IWikiPage>;
  title: string;
}

export interface IWikiPage {
  wikiPageTitle: string;
}


export default class WikiPageViewerWebPart extends BaseClientSideWebPart<IWikiPageViewerWebPartProps> {

   /**
   * Event handler for clicking the Configure button on the Placeholder
   */
  private _onConfigure = (): void => {
    this.context.propertyPane.open();
  }

  public render(): void {

    const needsConfiguration: boolean = !this.properties.wikiPage.tryGetSource();
    const element: React.ReactElement<IWikiPageViewerProps> = React.createElement(
      WikiPageViewer,
      {
        needsConfiguration: needsConfiguration,
        wikiPage: this.properties.wikiPage,
        onConfigure: this._onConfigure,
        displayMode: this.displayMode,
        title: this.properties.title,
        updateProperty: (value: string): void => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'wikiPage': {
        dynamicPropertyType: 'object'
      }
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      // pages: [
      //   {
      //     header: {
      //       description: strings.PropertyPaneDescription
      //     },
      //     groups: [
      //       {
      //         groupName: strings.BasicGroupName,
      //         groupFields: [
      //           PropertyPaneTextField('description', {
      //             label: strings.DescriptionFieldLabel
      //           })
      //         ]
      //       }
      //     ]
      //   }
      // ]
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Wiki Page',
                  fields: [
                    PropertyPaneDynamicField('wikiPage', {
                      label: 'wikiPage source'
                    })
                  ]
                  // sharedConfiguration: {
                  //   depth: DynamicDataSharedDepth.Property
                  // }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
