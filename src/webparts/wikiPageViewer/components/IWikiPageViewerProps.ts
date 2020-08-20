import { DynamicProperty } from "@microsoft/sp-component-base";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IWikiPage {
  wikiPageTitle: string;
}

export interface IWikiPageViewerProps {
  /**
   * Web part display mode. Used for inline editing of the web part title
   */
  displayMode: DisplayMode;
  /**
   * The title of the web part
   */
  title: string;
  /**
   * The currently selected event
   */
  wikiPage: DynamicProperty<IWikiPage>;
  /**
   * Determines if the web part has been connected to a dynamic data source or
   * not
   */
  needsConfiguration: boolean;
  /**
   * Event handler for clicking the Configure button on the Placeholder
   */
  onConfigure: () => void;
  /**
   * Event handler after updating the web part title
   */
  updateProperty: (value: string) => void;
}
