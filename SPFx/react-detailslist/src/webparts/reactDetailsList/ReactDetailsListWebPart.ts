import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactDetailsListWebPartStrings';
import {ReactDetailsList} from './components/ReactDetailsList';
import { IReactDetailsListProps } from './components/IReactDetailsListProps';

export interface IReactDetailsListWebPartProps {
  description: string;
}

export default class ReactDetailsListWebPart extends BaseClientSideWebPart<IReactDetailsListWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactDetailsListProps> = React.createElement(
      ReactDetailsList,
      {
        //description: this.properties.description
        context: this.context
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
