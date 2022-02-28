import * as React from 'react';
import { IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState } from '.';
import { SharePointServiceProvider } from '../../../api';
import { InformationType } from '../../../enums';
import { PagedUserInformation } from '../../../models';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartStrings';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import InfiniteScroll from 'react-infinite-scroller';
import { DefaultButton, DocumentCard, Image, Text } from '@fluentui/react';
import { LivePersona } from '@pnp/spfx-controls-react/lib/controls/LivePersona';
import { UserProfileInformation } from '../../../constants';
import styles from './BirthdaysWorkAnniversariesNewCollaboratorsMoreResults.module.scss'
import { DateHelper } from '../../../helpers';

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
    this.setBirthdaysFilter = this.setBirthdaysFilter.bind(this);
    this.setWorkAnniversariesFilter = this.setWorkAnniversariesFilter.bind(this);
    this.setNewCollaboratorsFilter = this.setNewCollaboratorsFilter.bind(this);
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

      const pagedUsers: PagedUserInformation = await this.sharePointServiceProvider.getPagedAnniversariesOrNewCollaborators(informationType,this.props.numberOfItemsPerPage, null);

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

      const pagedUsers: PagedUserInformation = await this.sharePointServiceProvider.getPagedAnniversariesOrNewCollaborators(informationType, this.props.numberOfItemsPerPage, this.state.nextPageUrl);

      if (pagedUsers.users != null && pagedUsers.users.length > 0) {
        this.setState({
          users: pagedUsers.users,
          nextPageUrl: pagedUsers.nextPageUrl
        });
      }
    }
  }

  private setBirthdaysFilter(): void {
    this.setState({
      informationType: InformationType.Birthdays
    });
  }

  private setWorkAnniversariesFilter(): void {
    this.setState({
      informationType: InformationType.WorkAnniversaries
    });
  }

  private setNewCollaboratorsFilter(): void {
    this.setState({
      informationType: InformationType.NewCollaborators
    });
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.users !== null && this.state.users.length > 0) {
        return (
          <>
            {strings.FilterByLabel}
            <button onClick={this.setBirthdaysFilter} className={styles.filterButton}>Birthdays</button>
            <button onClick={this.setWorkAnniversariesFilter} className={styles.filterButton}>Work Anniversaries</button>
            <button onClick={this.setNewCollaboratorsFilter} className={styles.filterButton}>New Collaborators</button>
            <InfiniteScroll
              loadMore={this.nextPage}
              hasMore={this.state.users !== null ? this.state.nextPageUrl.length > 0 : false}
              loader={<h4 className="loader" key={0}>Loading ...</h4>}
            >
              <>
                {
                  this.state.users !== null && this.state.users.map(user =>
                    <DocumentCard>
                        <div className={styles.persona}>
                          <LivePersona serviceScope={this.context.serviceScope} upn={user.Email}
                            template={
                              <>
                                <img className={styles.roundedImage} src={`${UserProfileInformation.profilePictureUrlPrefix + user.Email}`} />
                              </>
                            }
                          />
                        </div>
                        <div className={styles.title}>{user.Title}</div>
                        <div className={styles.date}>{DateHelper.getUserFormattedDate(user, this.state.informationType, strings.TodayLabel)}</div>
                    </DocumentCard>
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