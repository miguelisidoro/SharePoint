import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactBirthDaysAndNewHiresWebPartStrings';
import ReactBirthDaysAndNewHires from './components/ReactBirthDaysAndNewHires';
import { IReactBirthDaysAndNewHiresProps } from './components/IReactBirthDaysAndNewHiresProps';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

export interface IReactBirthDaysAndNewHiresWebPartProps {
  personalInformationListUrl: string;
  numberOfItemsToShow: number;
  typeOfInformation: string;
  birthDayNumberOfUpcomingDays: number;
  hireDateUpcomingDays: number;
}

export default class ReactBirthDaysAndNewHiresWebPart extends BaseClientSideWebPart<IReactBirthDaysAndNewHiresWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactBirthDaysAndNewHiresProps> = React.createElement(
      ReactBirthDaysAndNewHires,
      {
        personalInformationListUrl: this.properties.personalInformationListUrl,
        numberOfItemsToShow: this.properties.numberOfItemsToShow,
        birthDayNumberOfUpcomingDays: this.properties.birthDayNumberOfUpcomingDays,
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
                }),
                PropertyFieldNumber("birthDayNumberOfUpcomingDays", {
                  key: "birthDayNumberOfUpcomingDays",
                  label: strings.NumberUpComingDaysLabel,
                  description: strings.NumberUpComingDaysLabel,
                  value: this.properties.numberUpcomingDays,
                  maxValue: 10,
                  minValue: 5,
                  disabled: false
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
