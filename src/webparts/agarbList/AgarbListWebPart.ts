import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneSlider,
  PropertyPaneButton
} from "@microsoft/sp-property-pane";
import { escape } from "@microsoft/sp-lodash-subset";

import * as strings from "AgarbListWebPartStrings";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from "@microsoft/sp-http";

import AgarbList from "./components/AgarbList";
import { IAgarbListProps } from "./components/IAgarbListProps";
import IListItem from "./components/IListItem";

export interface IAgarbListWebPartProps {
  siteURL: string;
  top: number;
  ODataFilter: string;
  listName: string;
  newListName: string;
}

export default class AgarbListWebPart extends BaseClientSideWebPart<
  IAgarbListWebPartProps
> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private items: IListItem[] = [];

  protected onAfterDeserialize(
    deserializedObject: any,
    dataVersion: Version
  ): IAgarbListWebPartProps {
    let uniqueListName = new Date().valueOf()+"List";
    return {
      newListName: uniqueListName,
      ODataFilter: "",
      listName: "",
      siteURL: "",
      top: 5
    };
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
                resolve("");
                this.lists = [];
                this.context.propertyPane.refresh();
                this.showLists();
                this.render();
                return;
              } else if (response.status === 404) {
                resolve(
                  `List '${escape(value)}' doesn't exist in the current site`
                );
                setTimeout(() => {
                  this.properties.listName = value;
                  this.lists = [];
                  this.listsDropdownDisabled = true;
                  this.context.propertyPane.refresh();
                }, 1000);
                return;
              } else {
                resolve(`Error: ${response.statusText}. Please try again`);
                setTimeout(() => {
                  this.properties.listName = value;
                  this.lists = [];
                  this.listsDropdownDisabled = true;
                  this.context.propertyPane.refresh();
                }, 1000);
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
      if (this.items.length >= newValue) {
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
            )}')/Items?$top=${this.properties.top}&$select=${escape(filter)}`,
            SPHttpClient.configurations.v1
          )
          .then(response => response.json())
          .then(response => {
            resolve(
              response.value.map(data => {
                return {
                  ID: data.ID,
                  Title: data.Title,
                  Modified: data.Modified,
                  ModifiedBy: data.EditorId
                };
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

  protected showItems() {
    this.loadItems().then((items: IListItem[]) => {
      this.items = items;
      this.render();
    });
  }

  protected validateNewListTextField(value: string): string {
    console.log(this.properties.listName);
    if (this.properties.listName)
      return "You can't create new list if you've alredy picked one";
    else if (value === null || value.trim().length === 0) {
      return "Provide a new list name";
    }

    if (value.length > 40) {
      return "List name should not be longer than 40 characters";
    }

    return "";
  }

  protected createList() {
    if (this.properties.newListName == "") return;
    const options: ISPHttpClientOptions = {
      body: JSON.stringify({
        Title: this.properties.newListName,
        BaseTemplate: 100
      })
    };
    return new Promise<number>(
      (resolve: (result: number) => void, reject: (error: any) => void) => {
        this.context.spHttpClient
          .post(
            `https://agarb.sharepoint.com${escape(
              this.properties.siteURL
            )}/_api/web/lists`,
            SPHttpClient.configurations.v1,
            options
          )
          .then(response => response.json())
          .then(result => {
            this.properties.listName = this.properties.newListName;
            this.properties.newListName = "";
            this.showItems();
            resolve(0);
          })
          .catch(error => {
            console.log(error);
            reject(error);
          });
      }
    );
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
                PropertyPaneTextField("newListName", {
                  label: strings.newListNameLabel,
                  disabled: this.properties.listName != "",
                  onGetErrorMessage: this.validateNewListTextField.bind(this),
                  deferredValidationTime: 500
                }),
                PropertyPaneButton("createList", {
                  text: strings.CreateButtonText,
                  disabled: this.properties.listName != "",
                  onClick: this.createList.bind(this)
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
