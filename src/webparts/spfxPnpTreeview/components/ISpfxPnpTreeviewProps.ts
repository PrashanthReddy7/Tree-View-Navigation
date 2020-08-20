import { DisplayMode } from "@microsoft/sp-core-library";

export interface IWikiPage {
  wikiPageTitle: string;
}

export interface ISpfxPnpTreeviewProps {
  //description: string;  
  //_wikiPageTitleSelected:any | null;
  //displayMode: DisplayMode;
  /**
   * Web part display mode. Used for inline editing of the web part title
   */
  displayMode: DisplayMode;
  onWikiPageSelected: (wikiPage: IWikiPage) => void;
  title: string; 
  siteUrl: string;
  //context: any | null;
  updateProperty: (value: string) => void;
}
