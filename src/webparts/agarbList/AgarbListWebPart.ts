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
  IPropertyPaneDropdownOption,
  PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import styles from "./components/AgarbList.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";

import * as strings from "AgarbListWebPartStrings";
import MockHttpClient from "./MockHttpClient";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

import AgarbList from "./components/AgarbList";
import { IAgarbListProps } from "./components/IAgarbListProps";
import IListItem from "./components/IListItem";

export interface IAgarbListWebPartProps {
  siteURL: string;
  top: number;
  ODataFilter: string;
  listName: string;
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
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private items: IListItem[] = [];

  public render(): void {
    const element: React.ReactElement<IAgarbListProps> = React.createElement(
      AgarbList,
      {
        items:this.items
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private validateURL(value: string): Promise<string> {
    return new Promise<string>(
      (
        resolve: (validationErrorMessage: string) => void,
        reject: (error: any) => void
      ): void => {
        if (value === null || value.length === 0)
          value = this.context.pageContext.web.serverRelativeUrl;

        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com${escape(value)}`,
            SPHttpClient.configurations.v1
          )
          .then(
            (response: SPHttpClientResponse): void => {
              if (response.ok) {
                this.showLists();
                resolve("");
                return;
              } else if (response.status === 404) {
                resolve(
                  `List '${escape(value)}' doesn't exist in the current site`
                );
                this.lists = [];
                this.listsDropdownDisabled = true;
                return;
              } else {
                this.lists = [];
                this.listsDropdownDisabled = true;
                resolve(`Error: ${response.statusText}. Please try again`);
                return;
              }
            }
          )
          .catch(
            (error: any): void => {
              resolve(error);
            }
          );
      }
    );
  }

  protected onDispose(): void {
    this.properties.listName = "";
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ) {
    if (propertyPath === "listName" && newValue)
      this.showItems();
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com${escape(this.properties.siteURL)}/_api/web/lists?$filter=Hidden eq false`,
            SPHttpClient.configurations.v1
          )
          .then(response => response.json())
          .then(response => {
            resolve(
              response.value.map(option => {
                return { key: option.Title, text: option.Title };
              })
            );
          })
          .catch(error => {
            console.log(error);
            reject(error);
          });
      }
    );
  }

  protected showLists(): void {
    this.listsDropdownDisabled = !this.lists;

    if(this.lists) {
      this.listsDropdownDisabled = true;
      this.context.propertyPane.refresh();
    }

    this.loadLists().then(
      (listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.render();
      }
    );
  }

  private loadItems(): Promise<IListItem[]> {
    return new Promise<IListItem[]>(
      (
        resolve: (options: IListItem[]) => void,
        reject: (error: any) => void
      ) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com${escape(this.properties.siteURL)}/_api/lists/getByTitle('${escape(this.properties.listName)}')/Items`,
            SPHttpClient.configurations.v1
          )
          .then(response => response.json())
          .then(response => {
            resolve(response.value.map(data => {return {
              "ID":data.ID,
              "Title": data.Title,
              "Modified": data.Modified,
              "ModifiedBy": data.EditorId,
            }}));
          })
          .catch(error => {
            console.log(error);
            reject(error);
          });
      }
    );
  }

  protected showItems() {
    this.loadItems().then((items: IListItem[]) => {
      this.items = items;
      this.render();
    });
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
                  label: strings.SiteURLLabel,
                  value: this.context.pageContext.site.serverRelativeUrl,
                  onGetErrorMessage: this.validateURL.bind(this),
                  deferredValidationTime: 500
                }),
                PropertyPaneDropdown("listName", {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneSlider("top", {
                  label: strings.SliderLabel,
                  min: 1,
                  max: 20,
                  value: 5
                }),
                PropertyPaneTextField("ODataFilter", {
                  // label: strings.DescriptionFieldLabel
                  label: strings.ODataFilter,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
