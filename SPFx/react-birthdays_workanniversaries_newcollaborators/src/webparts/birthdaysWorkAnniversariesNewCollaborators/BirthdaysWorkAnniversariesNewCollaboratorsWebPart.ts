import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';

import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsWebPartStrings';
import BirthdaysWorkAnniversariesNewCollaborators from './components/BirthdaysWorkAnniversariesNewCollaborators';
import { IBirthdaysWorkAnniversariesNewCollaboratorsProps } from './components/IBirthdaysWorkAnniversariesNewCollaboratorsProps';

export interface IBirthdaysWorkAnniversariesNewCollaboratorsWebPartProps {
  sharePointRelativeListUrl: string;
  informationType: string;
  numberOfItemsToShow: number;
  numberOfDaysToRetrieve: number;
  title: string;
}

export default class BirthdaysWorkAnniversariesNewCollaboratorsWebPart extends BaseClientSideWebPart<IBirthdaysWorkAnniversariesNewCollaboratorsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsProps> = React.createElement(
      BirthdaysWorkAnniversariesNewCollaborators,
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.PropertiesGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.WebPartTitleLabel
                }),
                PropertyPaneTextField('sharePointRelativeListUrl', {
                  label: strings.SharePointRelativeListUrlFieldLabel
                }),
                PropertyFieldNumber("numberOfItemsToShow", {
                  key: "numberOfItemsToShow",
                  label: strings.NumberOfItemsToShowLabel,
                  description: strings.NumberOfItemsToShowLabel,
                  value: this.properties.numberOfItemsToShow,
                  maxValue: 20,
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
                    { key: 'NewCollaborators', text: strings.NewCollaboratorsInformationTypeLabel }, 
                  ]})
              ]
            }
          ]
        }
      ]
    };
  }
}
