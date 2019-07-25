import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IWebPartPropertiesMetadata
} from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import styles from "./components/AgarbList.module.scss";
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from "AgarbListWebPartStrings";
import MockHttpClient from "./MockHttpClient";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

import AgarbList from "./components/AgarbList";
import { IAgarbListProps } from "./components/IAgarbListProps";

export interface IAgarbListWebPartProps {
  siteURL: string;
  lists: string[];
  top: number;
  ODataFilter: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class AgarbListWebPart extends BaseClientSideWebPart<
  IAgarbListWebPartProps
> {
  public render(): void {
    //   this.domElement.innerHTML = `
    // <div class="${ styles.agarbList }">
    //   <div class="${ styles.container }">
    //     <div class="${ styles.row }">
    //       <div class="${ styles.column }">
    //         <span class="${ styles.title }">Welcome to SharePoint!</span>
    //         <p class="${ styles.subTitle }">Customize SharePoint experiences using web parts.</p>
    //         <p class="${ styles.description }">${escape(this.properties.description)}</p>
    //         <p class="${ styles.description }">Loading from ${escape(this.context.pageContext.web.title)}</p>
    //         <a href="https://aka.ms/spfx" class="${ styles.button }">
    //           <span class="${ styles.label }">Learn more</span>
    //         </a>
    //       </div>
    //     </div>
    //     <div id="spListContainer" />
    //   </div>
    // </div>`;

    // this._renderListAsync();

    const element: React.ReactElement<IAgarbListProps> = React.createElement(
      AgarbList,
      {
        siteURL: this.properties.siteURL,
        lists: this.properties.lists,
        top: this.properties.top,
        ODataFilter: this.properties.ODataFilter
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get().then((data: ISPList[]) => {
      var listData: ISPLists = { value: data };
      return listData;
    }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = "";
    items.forEach((item: ISPList) => {
      html += `
    <ul class="${styles.list}">
      <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });

    const listContainer: Element = this.domElement.querySelector(
      "#spListContainer"
    );
    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then(response => {
        this._renderList(response.value);
      });
    } else if (
      Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint
    ) {
      this._getListData().then(response => {
        this._renderList(response.value);
      });
    }
  }

  private validateURL(value: string): Promise<string> {
    return new Promise<string>((resolve: (validationErrorMessage: string) => void, reject: (error: any) => void): void => {
      if (value === null ||
        value.length === 0) value = this.context.pageContext.web.serverRelativeUrl;

      this.context.spHttpClient.get(`https://agarb.sharepoint.com${escape(value)}`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse): void => {
          if (response.ok) {
            resolve('');
            return;
          }
          else if (response.status === 404) {
            resolve(`List '${escape(value)}' doesn't exist in the current site`);
            return;
          }
          else {
            resolve(`Error: ${response.statusText}. Please try again`);
            return;
          }
        })
        .catch((error: any): void => {
          resolve(error);
        });
    });
  }

  // protected get propertiesMetadata(): IWebPartPropertiesMetadata {
  //   return {
  //     'title': { isSearchablePlainText: true },
  //     'intro': { isHtmlString: true },
  //     'image': { isImageSource: true },
  //     'url': { isLink: true }
  //   };
  // }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField("siteURL", {
                  label: strings.siteURLLabel,
                  value: this.context.pageContext.site.serverRelativeUrl,
                  onGetErrorMessage: this.validateURL.bind(this),
                  deferredValidationTime: 500
                }),
                PropertyPaneDropdown("lists", {
                  label: "Lists",
                  options: [
                    { index: 0, key: 0, text: "List1" },
                    { index: 1, key: 1, text: "List2" },
                    { index: 2, key: 2, text: "List3" }
                  ],
                  selectedKey: 0
                }),
                PropertyPaneSlider("top", {
                  label: "Top",
                  min: 1,
                  max: 20,
                  value: 5
                }),
                PropertyPaneTextField("ODataFilter", {
                  // label: strings.DescriptionFieldLabel
                  label: "Odata filter"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
