import * as React from 'react';
import styles from './WikiPageTags.module.scss';
import { IWikiPageTagsProps } from './IWikiPageTagsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jquery from 'jquery';
import ReactWordcloud from 'react-wordcloud';
import { select } from 'd3-selection';
import 'd3-transition';
import { IWikiPage } from "../../../data/IWikiPage";

function getCallback(callback) {
  return function (word, event) {
    const isActive = callback !== 'onWordMouseOut';
    const element = event.target;
    const text = select(element);
    text
      .on('click', () => {
        if (isActive) {
          window.open(`https://basfce1.sharepoint.com/teams/WikiReaderPOC/_layouts/15/search.aspx/siteall?q=${word.text}`, '_blank');
        }
      })
      .transition()
      .attr('background', 'white')
      //.attr('font-size', isActive ? '150%' : '100%')
      .attr('text-decoration', isActive ? 'underline' : 'none');
  };
}

const callbacks = {
  //getWordColor: word => (word.value > 50 ? 'orange' : 'purple'),
  getWordTooltip: word =>
    //` Search Wiki pages with "${word.text}" appears ${word.value} times.`,
    ` Search Wiki pages with "${word.text}"`,
  onWordClick: getCallback('onWordClick'),
  onWordMouseOut: getCallback('onWordMouseOut'),
  onWordMouseOver: getCallback('onWordMouseOver'),
};

export interface IWikiPageTagsState {
  items: [
    {
      "KeyPhrases": ""
    }];
  FinalKeyPhrase: [{ text: "", value: 0 }];
}

export default class WikiPageTags extends React.Component<IWikiPageTagsProps, IWikiPageTagsState> {
  //constructor(props: IWikiPageTagsProps) {
  public constructor(props: IWikiPageTagsProps, state: IWikiPageTagsState) {
    super(props);
    this.state = {
      items: [
        {
          "KeyPhrases": ""
        }],
      FinalKeyPhrase: [{ text: "", value: 0 }]
    };

  }
  public componentDidMount() {
    var reactHandler = this;
    //var siteURL=${this.props.siteurl};
    debugger;
    jquery.ajax({
      url: "https://basfce1.sharepoint.com/teams/WikiReaderPOC/_api/web/lists/getbytitle('site pages')/items?$select=Tags",
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {        
        var resultSet = resultData.d.results;
        var taggedKeyWords = [];

        resultSet.forEach(element => {
          if (element != null && element != undefined && element.Tags != null && element.Tags != undefined
            && element.Tags.results != null && element.Tags.results != undefined) {
            var docKeyTags = element.Tags.results;
            docKeyTags.forEach(eachtag => {
              var label = eachtag.Label;
              var guid = eachtag.TermGuid;
              //jquery.merge(taggedKeyWords, label);
              taggedKeyWords.push(label);
            });
          }
        });

        var keyPhrasesWithCount = [];
        taggedKeyWords.forEach(element => {
          if (keyPhrasesWithCount.length > 0) {
            var found = keyPhrasesWithCount.filter(x => x.text === element).length;
            if (found > 0) {
              keyPhrasesWithCount.filter(x => x.text === element)[0].value = keyPhrasesWithCount.filter(x => x.text === element)[0].value + 1;
            }
            else {
              keyPhrasesWithCount.push({ text: element, value: 1 });
            }
          }
          else {
            keyPhrasesWithCount.push({ text: element, value: 1 });
          }

        });
        Array.prototype.push.apply(reactHandler.state.FinalKeyPhrase, keyPhrasesWithCount);
        reactHandler.setState({ FinalKeyPhrase: reactHandler.state.FinalKeyPhrase });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });

  }

  public render(): React.ReactElement<IWikiPageTagsProps> {
    // return (
    //   <div className={ styles.wikiPageTags }>
    //     <div className={ styles.container }>
    //       <div className={ styles.row }>
    //         <div className={ styles.column }>
    //           <span className={ styles.title }>Welcome to SharePoint!</span>
    //           <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
    //           <p className={ styles.description }>{escape(this.props.description)}</p>
    //           <a href="https://aka.ms/spfx" className={ styles.button }>
    //             <span className={ styles.label }>Learn more</span>
    //           </a>
    //         </div>
    //       </div>
    //     </div>
    //   </div>
    // );
    const { needsConfiguration, wikiPage, onConfigure, displayMode, title, updateProperty } = this.props;
    const eventData: IWikiPage | undefined = wikiPage.tryGetValue();
    debugger;
    return (
      <div style={{ backgroundColor: '#efefef', height: '100%', width: '100%' }}>
        <ReactWordcloud
          callbacks={callbacks}
          words={this.state.FinalKeyPhrase}
          //minSize={[400, 400]}
          //size = {[400, 400]}
          options={{
            fontFamily: 'courier new',
            fontSizes: [10, 35],
            fontStyle: 'normal',
            fontWeight: 'bold',
            colors: [
              '#1f77b4',
              '#ff7f0e',
              '#2ca02c',
              '#d62728',
              '#9467bd',
              '#8c564b',
            ],
            rotations: 1,
            rotationAngles: [0, 0],
          }} />
      </div>
    );
  }
}
