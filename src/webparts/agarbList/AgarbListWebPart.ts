import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import { escape } from "@microsoft/sp-lodash-subset";

import * as strings from "AgarbListWebPartStrings";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import AgarbList from "./components/AgarbList";
import { IAgarbListProps } from "./components/IAgarbListProps";
import IListItem from "./components/IListItem";
import { number } from "prop-types";

export interface IAgarbListWebPartProps {
  siteURL: string;
  top: number;
  ODataFilter: string;
  listName: string;
}

export default class AgarbListWebPart extends BaseClientSideWebPart<
  IAgarbListWebPartProps
> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private items: IListItem[] = [];

  // constructor() {
  //   super();
  //   this.properties.ODataFilter="";
  //   this.properties.listName="";
  //   this.properties.siteURL="";
  //   this.properties.top=5;
  // }
  protected onPropertyPaneConfigurationStart(): void {
    this.properties.ODataFilter = "";
    this.properties.listName = "";
    this.properties.siteURL = "";
    this.properties.top = 5;
    this.context.propertyPane.refresh();
  }

  public render(): void {
    const element: React.ReactElement<IAgarbListProps> = React.createElement(
      AgarbList,
      {
        items: this.items
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
    if (propertyPath === "listName" && newValue) this.showItems();

    if (propertyPath === "top" && this.items) {
      console.log(this.items.length);
      if (this.items.length >= newValue) {
        console.log("cut");
        this.items = this.items.slice(0, newValue);
      } else this.showItems();
    }

    if (propertyPath === "ODataFilter") this.showItems();
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.siteURL)
      this.properties.siteURL = this.context.pageContext.site.serverRelativeUrl;
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com${escape(
              this.properties.siteURL
            )}/_api/web/lists?$filter=Hidden eq false`,
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

    if (this.lists) {
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

  private validateODataFilter(value: string): string {
    if (
      value == "" ||
      value == "ID" ||
      value == "Title" ||
      value == "Modified" ||
      value == "ModifiedBy" ||
      value == "EditorId"
    ) {
      return "";
    } else {
      return "This field does not exist";
    }
  }

  private loadItems(): Promise<IListItem[]> {
    const filter =
      this.properties.ODataFilter === ""
        ? `ID,Title,Modified,EditorId`
        : this.properties.ODataFilter;
    return new Promise<IListItem[]>(
      (
        resolve: (options: IListItem[]) => void,
        reject: (error: any) => void
      ) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com${escape(
              this.properties.siteURL
            )}/_api/lists/getByTitle('${escape(
              this.properties.listName
            )}')/Items?$select=${escape(filter)}`,
            SPHttpClient.configurations.v1
          )
          .then(response => response.json())
          .then(response => {
            resolve(
              response.value
                .map(data => {
                  return {
                    ID: data.ID,
                    Title: data.Title,
                    Modified: data.Modified,
                    ModifiedBy: data.EditorId
                  };
                })
                .slice(0, this.properties.top)
            );
          })
          .catch(error => {
            console.log(error);
            reject(error);
          });
      }
    );
  }

  protected showItems() {
    console.log("showing items");
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
                  label: strings.ODataFilter,
                  onGetErrorMessage: this.validateODataFilter.bind(this),
                  deferredValidationTime: 500
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
