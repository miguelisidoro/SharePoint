import * as React from 'react';
import { IBirthdaysWorkAnniversariesNewCollaboratorsProps } from './IBirthdaysWorkAnniversariesNewCollaboratorsProps';

import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsWebPartStrings';

import {
  Link,
  Persona,
  PersonaSize,
  IPersonaProps
} from '@fluentui/react';
import { LivePersona } from "@pnp/spfx-controls-react/lib/controls/LivePersona";

import { IBirthdaysWorkAnniversariesNewCollaboratorsState } from './IBirthdaysWorkAnniversariesNewCollaboratorsState';
import { SharePointServiceProvider } from '@app/api';
import { UserInformation } from '@app/models';
import { InformationType } from '@app/enums';
import { PersonaInformationMapper } from '@app/mappers';

const personaProps: IPersonaProps = {
  size: PersonaSize.size48,
  styles: {
    root: {
      width: 325,
      margin: 5
    }
  }
};
export default class BirthdaysWorkAnniversariesNewCollaborators extends React.Component<IBirthdaysWorkAnniversariesNewCollaboratorsProps, IBirthdaysWorkAnniversariesNewCollaboratorsState> {

  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props: IBirthdaysWorkAnniversariesNewCollaboratorsProps) {
    super(props);

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context,
      this.props.sharePointRelativeListUrl,
      this.props.numberOfDaysToRetrieve);

    this.state = {
      users: null
    };

    this.loadUsers = this.loadUsers.bind(this);
    this.isWebPartConfigured = this.isWebPartConfigured.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadUsers();
  }

  /// Checkes if web part is properly configured
  private isWebPartConfigured(): boolean {
    const isWebPartConfigured = (this.props.sharePointRelativeListUrl !== null && this.props.sharePointRelativeListUrl !== undefined
      && this.props.showMoreUrl !== null && this.props.showMoreUrl !== undefined
      && this.props.numberOfItemsToShow !== null && this.props.numberOfItemsToShow !== undefined
      && this.props.numberOfDaysToRetrieve !== null && this.props.numberOfDaysToRetrieve !== undefined
      && this.props.informationType !== null && this.props.informationType !== undefined);

    return isWebPartConfigured;
  }

  // Loads users from SharePoint
  private async loadUsers() {
    if (this.isWebPartConfigured()) {

      let users: UserInformation[] = [];

      const informationType: InformationType = InformationType[this.props.informationType];

      users = await this.sharePointServiceProvider.getAnniversariesOrNewCollaborators(informationType, this.props.numberOfItemsToShow);

      if (users !== null && users.length > 0) {
        const usersPersonInformation = PersonaInformationMapper.mapToPersonaInformations(users, informationType);

        this.setState({
          users: usersPersonInformation
        });
      }
    }
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.users !== null && this.state.users.length > 0) {
        return (
          <>
            <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateTitleProperty} />
            {
              this.state.users !== null && this.state.users.map(user =>
                <LivePersona key={user.userPrincipalName} upn={user.userPrincipalName}
                  template={
                    <>
                      <Persona {...user} {...personaProps} />
                    </>
                  }
                  serviceScope={this.props.context.serviceScope} />
              )
            }
            <Link href={this.props.showMoreUrl}>
              {strings.ShowMoreLabel}
            </Link>
          </>
        );
      } else {
        let noUsersMessage;
        const informationType: InformationType = InformationType[this.props.informationType];
        if (informationType === InformationType.Birthdays) {
          noUsersMessage = strings.NoBirthdaysLabel;
        } else if (informationType === InformationType.WorkAnniversaries) {
          noUsersMessage = strings.NoWorkAnniversariesLabel;
        } else {
          noUsersMessage = strings.NoNewHiresLabel;
        }
        return (<div>
          <WebPartTitle displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateTitleProperty} />
          {noUsersMessage}
        </div>);
      }
    } else {
      return (<div>
        {strings.WebPartConfigurationMissing}
      </div>);
    }
  }
}
