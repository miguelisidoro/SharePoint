import * as React from 'react';
import { IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState } from '.';
import { SharePointServiceProvider } from '../../../api';
import { InformationDisplayType, InformationType } from '../../../enums';
import { PagedUserInformation, UserInformation } from '../../../models';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartStrings';
import { PersonaInformationMapper } from '../../../mappers';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import InfiniteScroll from 'react-infinite-scroller';
import { IPersonaProps, Persona, PersonaSize } from '@fluentui/react';
import { LivePersona } from '@pnp/spfx-controls-react/lib/controls/LivePersona';

const personaProps: IPersonaProps = {
  size: PersonaSize.size48,
  styles: {
    root: {
      width: 325,
      margin: 5,
    },
  },
};
export default class BirthdaysWorkAnniversariesNewCollaboratorsMoreResults extends React.Component<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState> {

  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props) {
    super(props);

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context,
      this.props.sharePointRelativeListUrl,
      this.props.numberOfDaysToRetrieveForBirthdays);

    this.state = {
      users: null,
      nextPageUrl: null,
      informationType: InformationType.Birthdays,
    };

    this.loadUsers = this.loadUsers.bind(this);
    this.isWebPartConfigured = this.isWebPartConfigured.bind(this);
    this.nextPage = this.nextPage.bind(this);
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

      const informationType: InformationType = this.state.informationType;

      const pagedUsers: PagedUserInformation = await this.sharePointServiceProvider.getAnniversariesOrNewCollaborators(informationType, InformationDisplayType.MoreResults, this.props.numberOfItemsPerPage, null);

      if (pagedUsers.users != null && pagedUsers.users.length > 0) {
        this.setState({
          users: pagedUsers.users,
          nextPageUrl: pagedUsers.nextPageUrl
        });
      }
    }
  }

  private async nextPage() {
    if (this.isWebPartConfigured()) {

      const informationType: InformationType = this.state.informationType;

      const pagedUsers: PagedUserInformation = await this.sharePointServiceProvider.getAnniversariesOrNewCollaborators(informationType, InformationDisplayType.MoreResults, this.props.numberOfItemsPerPage, this.state.nextPageUrl);

      if (pagedUsers.users != null && pagedUsers.users.length > 0) {
        this.setState({
          users: pagedUsers.users,
          nextPageUrl: pagedUsers.nextPageUrl
        });
      }
    }
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.users !== null && this.state.users.length > 0) {
        return (
          <>
            <InfiniteScroll
              loadMore={this.nextPage}
              hasMore={this.state.users !== null ? this.state.nextPageUrl.length > 0 : false}
              loader={<h4 className="loader" key={0}>Loading ...</h4>}
            >
              <>
                {
                  this.state.users !== null && this.state.users.map(user =>
                    <LivePersona serviceScope={this.context.serviceScope} upn={user.Email}
                      template={
                        <>
                          <Persona {...user} {...personaProps} />
                        </>
                      }
                    />
                  )
                }
              </>
            </InfiniteScroll>
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