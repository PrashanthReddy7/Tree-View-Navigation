import { DisplayMode } from "@microsoft/sp-core-library";
import { DynamicProperty } from "@microsoft/sp-component-base";
import { IWikiPage } from "../../../data/IWikiPage";

export interface IWikiPageTagsProps {
  description: string;
   /**
   * Web part display mode. Used for inline editing of the web part title
   */
  displayMode: DisplayMode;
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
   * The title of the web part
   */
  title: string;
  /**
   * Event handler after updating the web part title
   */
  updateProperty: (value: string) => void;

}
