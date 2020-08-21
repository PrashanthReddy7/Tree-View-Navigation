import * as React from 'react';
import styles from './SpfxPnpTreeview.module.scss';
import { ISpfxPnpTreeviewProps } from './ISpfxPnpTreeviewProps';
import { ISpfxPnpTreeviewState } from './ISpfxPnpTreeviewState';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode } from "@pnp/spfx-controls-react/lib/TreeView";
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
//import ReactDOM from 'react-dom'

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IContextInfo } from "@pnp/sp/sites";

var treearr: ITreeItem[] = [];
var spDomainUrl = "";

function filterId(items, parentId) {
  var resultFound;
  var isReturned = false;
  resultFound = items.filter(function (value) { return value.label == parentId; });
  if (resultFound != undefined && resultFound != null && resultFound.length > 0) {
    isReturned = true;
    return resultFound;
  }
  if (resultFound.length == 0) {
    // Check in Childreen Elements
    items.forEach(element => {
      if (element.children != undefined && element.children != null && element.children.length > 0) {
        var filteredItem = filterId(element.children, parentId);
        if (filteredItem != undefined && filteredItem != null && filteredItem.length > 0) {
          resultFound = filteredItem;
          return filteredItem;
        }
      }
    });
    return resultFound;
  }
}

function addChild(curruntItem, allItems) {
  var oldItem;
  if (curruntItem["NavigationParent"] == null) {
    const tree: ITreeItem = {
      key: curruntItem.Id,
      label: curruntItem["Title"],
      data: spDomainUrl + curruntItem["FileRef"],
      children: []
    };
    treearr.push(tree);
  }
  else {

    var navigateParentString = curruntItem["NavigationParent"];
    var treecol: Array<ITreeItem> = filterId(treearr, navigateParentString);
    if (treecol == undefined || treecol == null || treecol.length == 0) {
      //Read missing parent Wiki Page
      var getMissingItem = allItems.filter(function (value) { return value["Title"] == navigateParentString; });
      if (getMissingItem.length != 0) {
        oldItem = curruntItem;
        //Include missing Item into Tree
        addChild(getMissingItem[0], allItems);
        //Include Previous Item also into tree
        addChild(oldItem, allItems);
      }
    }
    else if (treecol != undefined && treecol != null && treecol.length != 0) {
      const tree: ITreeItem = {
        key: curruntItem.Id,
        label: curruntItem["Title"],
        data: spDomainUrl + curruntItem["FileRef"],
        children: []
      };
      treecol[0].children.push(tree);
    }
  }
}

export default class SpfxPnpTreeview extends React.Component<ISpfxPnpTreeviewProps, ISpfxPnpTreeviewState> {
  constructor(props: ISpfxPnpTreeviewProps) {
    super(props);

    // sp.setup({
    //   spfxContext: this.props.context
    // });

    this.state = {
      TreeLinks: []
      //selectedWikiPageTitle: ""
    };
  }

  private _getSelection = (items: any[]): void => {
    // since the list allows selecting only one item, pick the first selected
    // event and pass to the event handler specified through component
    // properties
    //debugger;
    var selectedWikiTitle=items[0].label;
    this.props.onWikiPageSelected(selectedWikiTitle);
  }

  public componentWillMount() {
    //debugger;
    //this.props._wikiPageTitleSelected({ wikiPageTitle: this.state.selectedWikiPageTitle });
    this._getLinks();
    this.setState({ TreeLinks: treearr });
  }

  private async _getLinks() {
    //debugger;
    //const allItems: any[] = await sp.web.lists.getByTitle("Site Pages").items.getAll();
    //const oContext: IContextInfo = await sp.site.getContextInfo();
    const allItems: any[] = await sp.web.lists.getByTitle("Site Pages").items.select("Title", "NavigationParent", "ID", "NavigationOrder", "FileRef").orderBy("NavigationOrder", true).get();
    allItems.forEach(function (v, i) {

      //Check If item already added to tree array
      var curruntWikiTitle = v["Title"];
      var treecol: Array<ITreeItem> = filterId(treearr, curruntWikiTitle);
      if (treecol.length == 0) {
        addChild(v, allItems);
      }
      //console.log(v);
    });
    console.log(treearr);
    this.setState({ TreeLinks: treearr });
  }

  public render(): React.ReactElement<ISpfxPnpTreeviewProps> {
    return (
      <div className={styles.spfxPnpTreeview}>
        <TreeView
          items={this.state.TreeLinks}
          defaultExpanded={false}
          selectionMode={TreeViewSelectionMode.Single}
          selectChildrenIfParentSelected={false}
          showCheckboxes={false}
          treeItemActionsDisplayMode={TreeItemActionsDisplayMode.ContextualMenu}
          onSelect={this._getSelection}
          onExpandCollapse={this.onTreeItemExpandCollapse}
          //onRenderItem={this.renderCustomTreeItem} 
          />
      </div>
    );
  }

  private onTreeItemSelect(items: ITreeItem[]) {
    //debugger;
    console.log("Items selected: ", items);    
    //this.setState({ selectedWikiPageTitle: items[0].label });
    //this.props.._wikiPageTitleSelected({ wikiPageTitle: items });
  }

  private onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item);
  }

  private renderCustomTreeItem(item: ITreeItem): JSX.Element {
    return (
      <span>
        <a href={item.data} target={'_blank'}>
          {item.label}
        </a>
      </span>
    );
  }




}
