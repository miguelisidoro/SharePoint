import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactFileDownloadWebPartStrings';
import ReactFileDownload from './components/ReactFileDownload';
import { IReactFileDownloadProps } from './components/IReactFileDownloadProps';
import { sp } from "@pnp/sp";
//import "@pnp/sp/profiles";
import '@pnp/sp/webs';
import '@pnp/sp/site-users';

export interface IReactFileDownloadWebPartProps {
  description: string;
}

export default class ReactFileDownloadWebPart extends BaseClientSideWebPart<IReactFileDownloadWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactFileDownloadProps> = React.createElement(
      ReactFileDownload,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
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
