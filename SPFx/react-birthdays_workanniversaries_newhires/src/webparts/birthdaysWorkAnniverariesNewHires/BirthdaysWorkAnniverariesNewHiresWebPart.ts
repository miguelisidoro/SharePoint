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

import BirthdaysWorkAnniverariesNewHires from './components/BirthdaysWorkAnniverariesNewHires';
import { IBirthdaysWorkAnniverariesNewHiresProps } from './components/IBirthdaysWorkAnniverariesNewHiresProps';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

import { localizedStrings } from '../../loc/strings'

export interface IBirthdaysWorkAnniverariesNewHiresWebPartProps {
  sharePointRelativeListUrl: string;
  informationType: string;
  numberOfItemsToShow: number;
  numberOfDaysToRetrieve: number;
  title: string;
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
        displayMode: this.displayMode,
        title: this.properties.title,
        updateTitleProperty: (value: string) => {
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
            description: localizedStrings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: localizedStrings.PropertiesGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: localizedStrings.WebPartTitleLabel
                }),
                PropertyPaneTextField('sharePointRelativeListUrl', {
                  label: localizedStrings.SharePointRelativeListUrlFieldLabel
                }),
                PropertyFieldNumber("numberOfItemsToShow", {
                  key: "numberOfItemsToShow",
                  label: localizedStrings.NumberOfItemsToShowLabel,
                  description: localizedStrings.NumberOfItemsToShowLabel,
                  value: this.properties.numberOfItemsToShow,
                  maxValue: 10,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("numberOfDaysToRetrieve", {
                  key: "numberOfDaysToRetrieve",
                  label: localizedStrings.NumberOfDaysToRetrieveLabel,
                  description: localizedStrings.NumberOfDaysToRetrieveLabel,
                  value: this.properties.numberOfDaysToRetrieve,
                  maxValue: 90,
                  minValue: 1,
                  disabled: false
                }),
                PropertyPaneDropdown('informationType', {
                  label: localizedStrings.InformationTypeLabel,
                  options: [ 
                    { key: 'Birthdays', text: localizedStrings.BirthdaysInformationTypeLabel }, 
                    { key: 'WorkAnniversaries', text: localizedStrings.WorkAnniversariesInformationTypeLabel }, 
                    { key: 'NewHires', text: localizedStrings.NewHiresInformationTypeLabel }, 
                  ]})
              ]
            }
          ]
        }
      ]
    };
  }
}
