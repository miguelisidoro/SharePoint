import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BirthdaysWorkAnniverariesNewHiresWebPartStrings';
import BirthdaysWorkAnniverariesNewHires from './components/BirthdaysWorkAnniverariesNewHires';
import { IBirthdaysWorkAnniverariesNewHiresProps } from './components/IBirthdaysWorkAnniverariesNewHiresProps';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

export interface IBirthdaysWorkAnniverariesNewHiresWebPartProps {
  sharePointRelativeListUrl: string;
  informationType: string;
  numberOfItemsToShow: number;
  numberOfDaysToRetrieve: number;
}

export default class BirthdaysWorkAnniverariesNewHiresWebPart extends BaseClientSideWebPart<IBirthdaysWorkAnniverariesNewHiresWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBirthdaysWorkAnniverariesNewHiresProps> = React.createElement(
      BirthdaysWorkAnniverariesNewHires,
      {
        informationType: this.properties.informationType,
        numberOfItemsToShow: this.properties.numberOfItemsToShow,
        numberOfDaysToRetrieve: this.properties.numberOfDaysToRetrieve,//90,
        sharePointRelativeListUrl: this.properties.sharePointRelativeListUrl,
        context: this.context,
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
              groupName: strings.PropertiesGroupName,
              groupFields: [
                PropertyPaneTextField('sharePointRelativeListUrl', {
                  label: strings.SharePointRelativeListUrlFieldLabel
                }),
                PropertyFieldNumber("numberOfItemsToShow", {
                  key: "numberOfItemsToShow",
                  label: strings.NumberOfItemsToShowLabel,
                  description: strings.NumberOfItemsToShowLabel,
                  value: this.properties.numberOfItemsToShow,
                  maxValue: 10,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("numberOfDaysToRetrieve", {
                  key: "numberOfDaysToRetrieve",
                  label: strings.NumberOfDaysToRetrieveLabel,
                  description: strings.NumberOfDaysToRetrieveLabel,
                  value: this.properties.numberOfDaysToRetrieve,
                  maxValue: 90,
                  minValue: 1,
                  disabled: false
                }),
                PropertyPaneDropdown('informationType', {
                  label: strings.InformationTypeLabel,
                  options: [ 
                    { key: 'Birthdays', text: strings.BirthdaysInformationTypeLabel }, 
                    { key: 'WorkAnniversaries', text: strings.WorkAnniversariesInformationTypeLabel }, 
                    { key: 'NewHires', text: strings.NewHiresInformationTypeLabel }, 
                  ]})
              ]
            }
          ]
        }
      ]
    };
  }
}
