import * as React from 'react';
import { IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState } from '.';
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
import { SharePointServiceProvider } from '../../../api';

const emailImage: string = require('.../../../assets/email.png');
const teamsCallImage: string = require('.../../../assets/teams_call.png');
const teamsChatImage: string = require('.../../../assets/teams_chat.png');

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
    this.loadMore = this.loadMore.bind(this);
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

      const pagedUsers: PagedUserInformation = await this.sharePointServiceProvider.getPagedAnniversariesOrNewCollaborators(informationType, this.props.numberOfItemsPerPage, null);

      if (pagedUsers.users != null && pagedUsers.users.length > 0) {
        this.setState({
          users: pagedUsers.users,
          nextPageUrl: pagedUsers.nextPageUrl
        });
      }
    }
  }

  private async loadMore() {
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
    }, () => { this.loadUsers() });
  }

  private setWorkAnniversariesFilter(): void {
    this.setState({
      informationType: InformationType.WorkAnniversaries
    }, () => { this.loadUsers() });
  }

  private setNewCollaboratorsFilter(): void {
    this.setState({
      informationType: InformationType.NewCollaborators
    }, () => { this.loadUsers() });
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.users !== null && this.state.users.length > 0) {
        return (
          <div className={styles.mainContainer}>
            {strings.FilterByLabel}
            <button onClick={this.setBirthdaysFilter} className={this.state.informationType === InformationType.Birthdays ? styles.filterButtonActive : styles.filterButton}>{strings.BirthdaysLabel}</button>
            <button onClick={this.setWorkAnniversariesFilter} className={this.state.informationType === InformationType.WorkAnniversaries ? styles.filterButtonActive : styles.filterButton}>{strings.WorkAnniversariesLabel}</button>
            <button onClick={this.setNewCollaboratorsFilter} className={this.state.informationType === InformationType.NewCollaborators ? styles.filterButtonActive : styles.filterButton}>{strings.NewCollaboratorsLabel}</button>
            <div className={styles.contentContainer}>
              {
                this.state.users !== null && this.state.users.map(user =>
                  <div className={styles.itemContainer}>
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
                    <div className={styles.bottomImagesContainer}>
                      <div className={styles.emailContainer}>
                        <a href={`mailto:${user.Email}`}><img alt={strings.EmailToText} title={strings.EmailToText} src={emailImage} className={styles.emailImage} /></a>
                      </div>
                      <div className={styles.teamsCallContainer}>
                        <a href={`https://teams.microsoft.com/l/call/0/0?users=${user.Email}`}><img alt={strings.TeamsCallText} title={strings.TeamsCallText} src={teamsCallImage} className={styles.teamsCallImage} /></a>
                      </div>
                      <div className={styles.teamsChatContainer}>
                        <a href={`https://teams.microsoft.com/l/chat/0/0?users=${user.Email}`}><img alt={strings.TeamsChatText} title={strings.TeamsChatText} src={teamsChatImage} className={styles.teamsChatImage} /></a>
                      </div>
                    </div>
                  </div>
                )
              }
            </div>
            <div className={styles.bottomContainer}>
              <button className={styles.loadMoreButton} onClick={this.loadMore}>{strings.LoadMoreLabel}</button>
            </div>
          </div>
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