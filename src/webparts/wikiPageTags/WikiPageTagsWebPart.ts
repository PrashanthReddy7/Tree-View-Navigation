import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDynamicFieldSet, PropertyPaneDynamicField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';

import * as strings from 'WikiPageTagsWebPartStrings';
import WikiPageTags from './components/WikiPageTags';
import { IWikiPageTagsProps } from './components/IWikiPageTagsProps';
import { IWikiPage } from "../../data/IWikiPage";
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IWikiPageTagsWebPartProps {
  description: string;
   /**
   * Event to show the details for
   */
  wikiPage: DynamicProperty<IWikiPage>;
  /**
   * Web part title
   */
  title: string;
}

export default class WikiPageTagsWebPart extends BaseClientSideWebPart<IWikiPageTagsWebPartProps> {

  /**
   * Event handler for clicking the Configure button on the Placeholder 
   */
  private _onConfigure = (): void => {
    this.context.propertyPane.open();
  }

  public render(): void {
    const needsConfiguration: boolean = !this.properties.wikiPage.tryGetSource();
    const element: React.ReactElement<IWikiPageTagsProps> = React.createElement(
      WikiPageTags,
      {
        description: this.properties.description,
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
    // return {
    //   pages: [
    //     {
    //       header: {
    //         description: strings.PropertyPaneDescription
    //       },
    //       groups: [
    //         {
    //           groupName: strings.BasicGroupName,
    //           groupFields: [
    //             PropertyPaneTextField('description', {
    //               label: strings.DescriptionFieldLabel
    //             })
    //           ]
    //         }
    //       ]
    //     }
    //   ]
    // };
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Select Wiki source',
                  fields: [
                    PropertyPaneDynamicField('wikiPage', {
                      label: 'Wiki source'
                    })
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
