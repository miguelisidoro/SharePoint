import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartStrings';
import BirthdaysWorkAnniversariesNewCollaboratorsMoreResults from './components/BirthdaysWorkAnniversariesNewCollaboratorsMoreResults';
import { IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps } from './components/IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

export interface IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartProps {
  title: string;
  sharePointRelativeListUrl: string;
  numberOfItemsPerPage: number;
  numberOfDaysToRetrieveForBirthdays: number;
  numberOfDaysToRetrieveForWorkAnniveraries: number;
  numberOfDaysToRetrieveForNewCollaborators: number;
}

export default class BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPart extends BaseClientSideWebPart<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps> = React.createElement(
      BirthdaysWorkAnniversariesNewCollaboratorsMoreResults,
      {
        numberOfItemsPerPage: this.properties.numberOfItemsPerPage,
        numberOfDaysToRetrieveForBirthdays: this.properties.numberOfDaysToRetrieveForBirthdays,
        numberOfDaysToRetrieveForWorkAnniveraries: 
        this.properties.numberOfDaysToRetrieveForWorkAnniveraries,
        numberOfDaysToRetrieveForNewCollaborators: 
        this.properties.numberOfDaysToRetrieveForNewCollaborators,
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
                PropertyFieldNumber("numberOfItemsPerPage", {
                  key: "numberOfDaysToRetrieveForBirthdays",
                  label: strings.NumberOfItemsPerPageLabel,
                  value: this.properties.numberOfItemsPerPage,
                  maxValue: 90,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("numberOfDaysToRetrieveForBirthdays", {
                  key: "numberOfDaysToRetrieveForBirthdays",
                  label: strings.NumberOfDaysToRetrieveLabelForBirthdays,
                  value: this.properties.numberOfDaysToRetrieveForBirthdays,
                  maxValue: 90,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("numberOfDaysToRetrieveForWorkAnniveraries", {
                  key: "numberOfDaysToRetrieveForWorkAnniveraries",
                  label: strings.NumberOfDaysToRetrieveLabelForWorkAnniversaries,
                  value: this.properties.numberOfDaysToRetrieveForWorkAnniveraries,
                  maxValue: 90,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("numberOfDaysToRetrieveForNewCollaborators", {
                  key: "numberOfDaysToRetrieveForNewCollaborators",
                  label: strings.NumberOfDaysToRetrieveLabelForNewCollaborators,
                  value: this.properties.numberOfDaysToRetrieveForNewCollaborators,
                  maxValue: 90,
                  minValue: 1,
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
