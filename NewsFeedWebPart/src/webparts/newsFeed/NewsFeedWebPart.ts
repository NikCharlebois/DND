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

export interface ISPNewsItems
{
  value: ISPNewsItem[];
}
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


import styles from './NewsFeedWebPart.module.scss';
import * as strings from 'NewsFeedWebPartStrings';

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

  private _getListData(): Promise<ISPNewsItems>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${ escape(this.properties.newsfeedListName)}')/Items`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) =>
        {
          debugger;
          return response.json();
        });
  }

  private _getSpecificItem(ID: string): Promise<ISPNewsItem>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${ escape(this.properties.newsfeedListName)}')/Items/GetByID('${ ID }')`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) =>
        {
          debugger;
          return response.json();
        });
  }

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
      
      var closePopupLink = document.createElement("a");
      closePopupLink.addEventListener("click", (e:Event) =>  this._hidePopup(item.ID));
      closePopupLink.className = styles.PopupCloseLink;
      closePopupLink.href = "#";
      closePopupLink.innerHTML = "X";
      popup.appendChild(closePopupLink);
      popupContent.appendChild(popup);
    }
  }

  private _hidePopup(ID):void
  {
    var remove = document.getElementById("divPopup" + ID);
    remove.style.display = "none";
    if (remove) remove.parentNode.removeChild(remove);
  }

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

  protected get dataVersion(): Version
  {
    return Version.parse('1.0');
  }

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
