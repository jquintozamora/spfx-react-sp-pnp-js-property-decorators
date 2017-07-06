import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "customBusinessObjectsPnPJsStrings";
import CustomBusinessObjectsPnPJs from "./components/CustomBusinessObjectsPnPJs";
import { ICustomBusinessObjectsPnPJsProps } from "./components/ICustomBusinessObjectsPnPJsProps";
import { ICustomBusinessObjectsPnPJsWebPartProps } from "./ICustomBusinessObjectsPnPJsWebPartProps";

import CodeSlider from "./components/CodeSlider/CodeSlider";
import { ICodeSliderProps } from "./components/CodeSlider/ICodeSliderProps";

import pnp, {
  Logger,
  ConsoleListener,
  LogLevel
} from "sp-pnp-js";

export default class CustomBusinessObjectsPnPJsWebPart extends BaseClientSideWebPart<ICustomBusinessObjectsPnPJsWebPartProps> {

  // // https://github.com/SharePoint/PnP-JS-Core/wiki/Using-sp-pnp-js-in-SharePoint-Framework
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // establish SPFx context
      pnp.setup({
        spfxContext: this.context
      });
      // subscribe a listener
      Logger.subscribe(new ConsoleListener());
      // set the active log level
      Logger.activeLogLevel = LogLevel.Warning;

    });
  }

  public render(): void {

    // Blog demo: https://blog.josequinto.com/2017/05/19/why-do-we-should-use-custom-business-objects-models-in-pnp-js-core/
    // const element: React.ReactElement<ICustomBusinessObjectsPnPJsProps> = React.createElement(
    //   CustomBusinessObjectsPnPJs,
    //   {
    //     description: this.properties.description
    //   }
    // );

    // PnP SIG Demo
    const element: React.ReactElement<ICodeSliderProps> = React.createElement(
      CodeSlider,
      {
        title: "PnP JS Core - Custom Business Objects DEMO!"
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
