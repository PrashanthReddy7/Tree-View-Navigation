import * as React from 'react';
import styles from './WikiPageViewer.module.scss';
import { IWikiPageViewerProps } from './IWikiPageViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import Iframe from 'react-iframe';

export interface IWikiPage {
  wikiPageTitle: string;
}


export default class WikiPageViewer extends React.Component<IWikiPageViewerProps, {}> {
  public render(): React.ReactElement<IWikiPageViewerProps> {
    const { displayMode, title, wikiPage, needsConfiguration, onConfigure, updateProperty } = this.props;
    const wikiData: IWikiPage | undefined = wikiPage.tryGetValue();
    const siteURL = "/teams/WikiReaderPOC/SitePages/";
    var pageUrl = '';
    if (wikiData != null && wikiData != undefined)
      pageUrl = siteURL + wikiData + ".aspx";
    //debugger;

    return (
      <div className={styles.wikiPageViewer}>
        <WebPartTitle displayMode={displayMode}
          title={title}
          updateProperty={updateProperty} />
        {needsConfiguration &&
          <Placeholder
            iconName='Edit'
            iconText='Configure your web part'
            description='Please configure the web part.'
            buttonLabel='Configure'
            onConfigure={onConfigure} />}
        {!needsConfiguration &&
          !wikiData &&
          <Placeholder
            iconName='CustomList'
            iconText='Wiki Page Content'
            description='Select wiki page' />}
        {!needsConfiguration &&
          wikiData &&
          <Iframe url={pageUrl}
            width="100%"
            height="800px"
            id="myPage"
            position="static" />}
      </div>
    );
  }
}
