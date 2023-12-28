/* eslint-disable no-unused-expressions */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import TestPart from './components/TestPart';
import { ITestPartProps } from './components/ITestPartProps';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { setup as pnpSetup } from "@pnp/common";

export interface ITestPartWebPartProps {
  description: string;
}

export default class TestPartWebPart extends BaseClientSideWebPart<ITestPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITestPartProps> = React.createElement(
      TestPart,
      {
        context: {
          spHttpClient: this.context.spHttpClient,
          pageContext: { web: { absoluteUrl: this.context.pageContext.web.absoluteUrl } }
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }


  public async onInit(): Promise<void> {
    try {
      false && SPComponentLoader.loadCss(
        "https://zlendoit.sharepoint.com/sites/SPTraining/SiteAssets/CSS/index.css"
      );
      pnpSetup({
        ie11: true,
        spfxContent: this.context
      });
    }
    catch (err) {
      console.log(err, "error to Load");

    }
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


}
