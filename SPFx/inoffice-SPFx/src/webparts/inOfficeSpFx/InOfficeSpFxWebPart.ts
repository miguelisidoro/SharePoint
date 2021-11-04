import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'InOfficeSpFxWebPartStrings';
import InOfficeSpFx from './components/InOfficeSpFxComponent/InOfficeSpFx';
import { IInOfficeSpFxProps } from './components/InOfficeSpFxComponent/IInOfficeSpFxProps';

export interface IInOfficeSpFxWebPartProps {
  title: string;
}

export default class InOfficeSpFxWebPart extends BaseClientSideWebPart<IInOfficeSpFxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IInOfficeSpFxProps> = React.createElement(
      InOfficeSpFx,
      {
        title: this.properties.title,
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
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
