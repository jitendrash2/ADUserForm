import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as React from "react";
import * as ReactDom from "react-dom";
import AdUserForm from "./components/AdUserForm";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

interface IAdUserFormWebPartProps {}

export default class AdUserFormWebPart extends BaseClientSideWebPart<IAdUserFormWebPartProps> {
  public render(): void {
    this.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        const element = React.createElement(AdUserForm, { graphClient: client });
        ReactDom.render(element, this.domElement);
      })
      .catch((error) => {
        console.error("Failed to initialize the Microsoft Graph client:", error);
      });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
