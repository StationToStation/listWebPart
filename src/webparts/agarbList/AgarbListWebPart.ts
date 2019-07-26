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

  public render(): void {
    const element: React.ReactElement<IAgarbListProps> = React.createElement(
      AgarbList,
      {
        siteURL: this.properties.siteURL,
        top: this.properties.top,
        ODataFilter: this.properties.ODataFilter,
        listName: this.properties.listName || ""
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
                this.showLists();
                return;
              } else if (response.status === 404) {
                resolve(
                  `List '${escape(value)}' doesn't exist in the current site`
                );
                return;
              } else {
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

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ) {
    console.log(propertyPath + ": " + oldValue + " -> " + newValue);
    // this.render();
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        fetch("https://agarb.sharepoint.com/sites/dev2/_api/web/lists", {
          headers: {
            accept: "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose"
          }
        })
          .then(resonse => resonse.json())
          .then(response => resolve(response.d.results.map(option => {return {key: option.Title, text: option.Title}})))
            //{response.d.results.map(option => {return {key: option.Id, text: option.Title}})})
          .catch(error => reject(error));
      }
    );
  }

  protected showLists(): void {
    this.listsDropdownDisabled = !this.lists;

    if (this.lists) {
      return;
    }

    // this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
    .then((listOptions: IPropertyPaneDropdownOption[]): void => {
      console.log(listOptions);
      this.lists = listOptions;
      this.listsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      // this.context.statusRenderer.clearLoadingIndicator(this.domElement);
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
                  label: strings.siteURLLabel,
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
