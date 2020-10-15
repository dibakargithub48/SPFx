import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "CrudoperationWebPartStrings";
import Crudoperation from "./components/Crudoperation";
import { ICrudoperationProps } from "./components/ICrudoperationProps";
import { sp } from "@pnp/sp/presets/all";

export interface ICrudoperationWebPartProps {
  description: string;
  listName: string;
}

export default class CrudoperationWebPart extends BaseClientSideWebPart<
  ICrudoperationWebPartProps
> {
  protected async onInit(): Promise<void> {
    await super.onInit();
    // other init code may be present
    sp.setup(this.context);
  }

  public render(): void {
    const element: React.ReactElement<ICrudoperationProps> = React.createElement(
      Crudoperation,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneTextField("listName", {
                  label: strings.ListNameFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
