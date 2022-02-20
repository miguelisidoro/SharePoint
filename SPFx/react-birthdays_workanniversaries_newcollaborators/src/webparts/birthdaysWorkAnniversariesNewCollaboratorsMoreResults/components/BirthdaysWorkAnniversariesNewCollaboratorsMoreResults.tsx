import * as React from 'react';
import { IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState } from '.';
import { SharePointServiceProvider } from '../../../api';
import { InformationDisplayType, InformationType } from '../../../enums';
import { UserInformation } from '../../../models';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartStrings';
import { PersonaInformationMapper } from '../../../mappers';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';

export default class BirthdaysWorkAnniversariesNewCollaboratorsMoreResults extends React.Component<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState> {

  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props) {
    super(props);

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context,
      this.props.sharePointRelativeListUrl,
      this.props.numberOfDaysToRetrieveForBirthdays);

    this.state = {
      allUsers: null,
      pagedUsers: null,
      usersToShow: null,
      informationType: InformationType.Birthdays,
    };

    this.loadUsers = this.loadUsers.bind(this);
    this.isWebPartConfigured = this.isWebPartConfigured.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadUsers();
  }

  /// Checkes if web part is properly configured
  private isWebPartConfigured(): boolean {
    const isWebPartConfigured = (this.props.sharePointRelativeListUrl != null && this.props.sharePointRelativeListUrl != undefined
      && this.props.numberOfDaysToRetrieveForBirthdays != null && this.props.numberOfDaysToRetrieveForBirthdays != undefined
      && this.props.numberOfDaysToRetrieveForNewCollaborators !== null && this.props.numberOfDaysToRetrieveForNewCollaborators !== undefined
      && this.props.numberOfDaysToRetrieveForWorkAnniveraries !== null && this.props.numberOfDaysToRetrieveForWorkAnniveraries !== undefined
      && this.props.numberOfItemsPerPage !== null && this.props.numberOfItemsPerPage !== undefined);

    return isWebPartConfigured;
  }

  // Loads users from SharePoint
  private async loadUsers() {
    if (this.isWebPartConfigured()) {

      let users: UserInformation[] = [];

      const informationType: InformationType = this.state.informationType;

      let numberOfDaysToRetrieve;

      if (informationType === InformationType.Birthdays) {
        numberOfDaysToRetrieve = this.props.numberOfDaysToRetrieveForBirthdays;
      }
      else if (informationType === InformationType.WorkAnniversaries) {
        numberOfDaysToRetrieve = this.props.numberOfDaysToRetrieveForWorkAnniveraries;
      }
      else //New Collaborators
      {
        numberOfDaysToRetrieve = this.props.numberOfDaysToRetrieveForNewCollaborators;
      }

      users = await this.sharePointServiceProvider.getAnniversariesOrNewCollaborators(informationType, InformationDisplayType.MoreResults, numberOfDaysToRetrieve);

      if (users != null && users.length > 0) {
        let usersPersonInformation = PersonaInformationMapper.mapToPersonaInformations(users, informationType);

        this.setState({
          allUsers: users,
        });
      }
    }
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.allUsers !== null && this.state.allUsers.length > 0) {
        return (
          <>
          </>
        );
      }
      else {
        let noUsersMessage;
        const informationType: InformationType = this.state.informationType;
        if (informationType === InformationType.Birthdays) {
          noUsersMessage = strings.NoBirthdaysLabel;
        }
        else if (informationType === InformationType.WorkAnniversaries) {
          noUsersMessage = strings.NoWorkAnniversariesLabel;
        }
        else {
          noUsersMessage = strings.NoNewHiresLabel;
        }
        return (<div>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateTitleProperty} />
          {noUsersMessage}
        </div>);
      }
    }
    else {
      return (<div>
        {strings.WebPartConfigurationMissing}
      </div>);
    }
  }
}