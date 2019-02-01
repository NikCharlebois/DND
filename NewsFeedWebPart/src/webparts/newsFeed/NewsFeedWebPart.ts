import
{
  Environment,
  Version,
  EnvironmentType
}
from '@microsoft/sp-core-library';

import
{
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
}
from '@microsoft/sp-webpart-base';

import
{
  escape
}
from '@microsoft/sp-lodash-subset';

// Defines an object that is an array of NewsFeed items;
export interface ISPNewsItems
{
  value: ISPNewsItem[];
}

// Defines what fields are expected to be part of a NewsFeed item'
export interface ISPNewsItem
{
  ID: string,
  TitleEnglish: string;
  TitleFrench: string;
  ContentEnglish: string;
  ContentFrench: string;
}

import MockHttpClient from './MockHttpClient';

import
{
  SPHttpClient,
  SPHttpClientResponse
}
from '@microsoft/sp-http';

// Import the CSS and locales;
import styles from './NewsFeedWebPart.module.scss';
import * as strings from 'NewsFeedWebPartStrings';

// This determines what properties needs to be exposed in the web part's properties pane;
export interface INewsFeedWebPartProps
{
  newsfeedListName: string;
  titleEnglishColumn: string;
  titleFrenchColumn: string;
  contentEnglishColumn: string;
  contentFrenchColumn: string;
}

export default class NewsFeedWebPart extends BaseClientSideWebPart<INewsFeedWebPartProps>
{
  /* This method mocks the retrieval of News item in the workbench only. It will be ignored when published to an actual SharePoint site. */
  private _getMockListData(): Promise<ISPNewsItems>
  {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => 
    {
        const listData: ISPNewsItems = 
        {
          value:
          [
            { ID: '1', TitleEnglish: '50/50 winners', TitleFrench: 'Gagnants du 50/50', ContentEnglish: 'Bob won 2$',ContentFrench: 'Bob a gagne 2$' },
            { ID: '2',  TitleEnglish: 'Email Outage', TitleFrench: 'Panne de Courriel', ContentEnglish: 'There is an ongoing email outage',ContentFrench: 'Il y a presentement une panne de courriel' },
            { ID: '5',  TitleEnglish: 'St-Patricks Party', TitleFrench: 'Fete de la St-Patrick', ContentEnglish: 'Buy your tickets today!',ContentFrench: "Achetez vos billets des aujourd'hui" }
          ]
        };
        return listData;
    }) as Promise<ISPNewsItems>;
  }

  // This Fakes the REST API that will be returning the content of a single list item based on its ID.
  private _getMockSpecificItem(ID: string): Promise<ISPNewsItem>
  {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => 
    {
        const listData: ISPNewsItem = 
        {
             ID: '1', 
             TitleEnglish: '50/50 winners', 
             TitleFrench: 'Gagnants du 50/50', 
             ContentEnglish: 'Bob won 2$',
             ContentFrench: 'Bob a gagne 2$'
        };
        return listData;
    }) as Promise<ISPNewsItem>;
  }

  // Makes a call to the SharePoint REST APIs to retrieve all items of a list based on its provided Title;
  private _getListData(): Promise<ISPNewsItems>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${ escape(this.properties.newsfeedListName)}')/Items`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) =>
        {
          debugger;
          return response.json();
        });
  }

  // Makes call to the SharePoint REST APIs to retrieve a single list item based on ID from the specified list;
  private _getSpecificItem(ID: string): Promise<ISPNewsItem>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${ escape(this.properties.newsfeedListName)}')/Items/GetByID('${ ID }')`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) =>
        {
          debugger;
          return response.json();
        });
  }

  // If we are running the local workbench (gulp serve), then use the Mock functions and data, otherwise use the official ones;
  private _getSpecificItemAsync(ID: string): void
  {
    if (Environment.type === EnvironmentType.Local)
    {
      this._getMockSpecificItem(ID).then((response) => 
      {
        this._renderItemDisplay(response);
      });
    }
    else
    {
      this._getSpecificItem(ID)
      .then((response) =>
      {
        this._renderItemDisplay(response);
      });
    }
  }

  // If we are running the local workbench (gulp serve), then use the Mock functions and data, otherwise use the official ones;
  private _renderListAsync(): void
  {
    if (Environment.type === EnvironmentType.Local)
    {
      this._getMockListData().then((response) => 
      {
        this._renderList(response.value);
      });
    }
    else
    {
      this._getListData()
      .then((response) =>
      {
        this._renderList(response.value);
      });
    }
  }

  // Renders a single item's content as a rounded Div located right under the item's title select by the user;
  private _renderItemDisplay(item: ISPNewsItem)
  {
    var existingPopup = document.getElementById("divPopup" + item.ID);
    if(null == existingPopup)
    {
      var popup = document.createElement("div");
      popup.id = "divPopup" + item.ID;
      popup.className = styles.NewsFeedPopup;
      popup.innerHTML = item[this.properties.contentEnglishColumn];
      var popupContent = document.getElementById("popupContent" + item.ID);
      
      // Define an X link to close the news item upon being clicked;
      var closePopupLink = document.createElement("a");
      closePopupLink.addEventListener("click", (e:Event) =>  this._hidePopup(item.ID));
      closePopupLink.className = styles.PopupCloseLink;
      closePopupLink.href = "#";
      closePopupLink.innerHTML = "X";
      popup.appendChild(closePopupLink);
      popupContent.appendChild(popup);
    }
  }

  // Hides the news item for which the associated X link was clicked;
  private _hidePopup(ID):void
  {
    var remove = document.getElementById("divPopup" + ID);
    remove.style.display = "none";
    if (remove) remove.parentNode.removeChild(remove);
  }

  // Retrieve all NewsFeed items from the specified list and render them as links based on their titles;
  private _renderList(items: ISPNewsItem[]): void
  {
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    items.forEach((item: ISPNewsItem) =>
    {
      var linkName = "lnkNewsFeedItem" + item.ID;
      var currentLink = document.createElement("a");
      currentLink.id = linkName;
      currentLink.href = "#";
      currentLink.className = styles.NewsItem;
      currentLink.addEventListener("click", (e:Event) =>  this._getSpecificItemAsync(item.ID));
      currentLink.innerHTML = item[this.properties.titleEnglishColumn];
      listContainer.appendChild(currentLink);
      var lineBreak = document.createElement("br");
      listContainer.appendChild(lineBreak);

      var popupContent = document.createElement("div");
      popupContent.id = "popupContent" + item.ID;
      listContainer.appendChild(popupContent);
    }); 
  }

  // Main method being executed when the web part loads. MAkes asynchrounous calls to retrieve the data;
  public render(): void
  {
    var currentUrl = window.location.href;
    this.domElement.innerHTML = `
      <div class="${ styles.newsFeed }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
            <div id="spListContainer" />
            </div>
          </div>
        </div>
      </div>`;
      this._renderListAsync();
  }

  // Displays the SPFx webpart's menu;
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration
  {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('newsfeedListName', {
                  label: strings.NewsFeedNameFieldLabel
                }),
                PropertyPaneTextField('titleEnglishColumn', {
                  label: strings.TitleEnglishColumn
                }),
                PropertyPaneTextField('titleFrenchColumn', {
                  label: strings.TitleFrenchColumn
                }),
                PropertyPaneTextField('contentEnglishColumn', {
                  label: strings.ContentEnglishColumn
                }),
                PropertyPaneTextField('contentFrenchColumn', {
                  label: strings.ContentFrenchColumn
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
