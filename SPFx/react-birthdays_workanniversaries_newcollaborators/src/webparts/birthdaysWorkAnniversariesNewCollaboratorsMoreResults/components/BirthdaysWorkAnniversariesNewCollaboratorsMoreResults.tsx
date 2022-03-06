import * as React from 'react';
import { IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps, IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsState } from '.';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsMoreResultsWebPartStrings';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { LivePersona } from '@pnp/spfx-controls-react/lib/controls/LivePersona';
import styles from './BirthdaysWorkAnniversariesNewCollaboratorsMoreResults.module.scss';
import { SharePointServiceProvider } from '@app/api';
import { InformationType } from '@app/enums';
import { PagedUserInformation } from '@app/models';
import { UserProfileInformation } from '@app/constants';
import { DateHelper } from '@app/helpers';
import { Dropdown, FontIcon, IDropdownOption, mergeStyles } from '@fluentui/react';

const emailImage: string = require('.../../../assets/email.png');
const teamsCallImage: string = require('.../../../assets/teams_call.png');
const teamsChatImage: string = require('.../../../assets/teams_chat.png');

const emailIconClass = mergeStyles({
  fontSize: 16,
  height: 16,
  width: 16,
  margin: '0 5px',
  color: 'black'
});

const teamsIconClass = mergeStyles({
  fontSize: 16,
  height: 16,
  width: 16,
  margin: '0 5px',
  color: '#4149B3'
});

const teamsCallIconClass = mergeStyles({
  fontSize: 16,
  height: 16,
  width: 16,
  margin: '0 5px',
  color: '#4149B3'
});

const filterDropDownOptions = [
  { key: 'Birthdays', text: 'Birthdays' },
  { key: 'WorkAnniversaries', text: 'Work Anniversaries' },
  { key: 'NewCollaborators', text: 'New Collaborators' },
];

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
    this.onFilterChange = this.onFilterChange.bind(this);
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

  private onFilterChange(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {
    const informationType: InformationType = InformationType[item.key];
    this.setState({
      informationType: informationType
    }, () => { this.loadUsers(); });
  };

  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsMoreResultsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.users !== null && this.state.users.length > 0) {
        return (
          <div className={styles.mainContainer}>
            <div>
              <div className={styles.fliterLabel}>
                {strings.FilterByLabel}
              </div>
              <div className={styles.filterDropdownContainer}>
                <Dropdown
                  onChange={this.onFilterChange}
                  placeholder="Select an option"
                  options={filterDropDownOptions}
                  className={styles.filterDropdown}
                />
              </div>
            </div>
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
                        <a href={`mailto:${user.Email}`}>
                          <FontIcon aria-label="Compass" iconName="MailSolid" className={emailIconClass} />
                        </a>
                      </div>
                      <div className={styles.teamsCallContainer}>
                        <a href={`https://teams.microsoft.com/l/call/0/0?users=${user.Email}`}>
                          <FontIcon aria-label="Compass" iconName="Phone" className={teamsCallIconClass} />
                        </a>
                      </div>
                      <div className={styles.teamsChatContainer}>
                        <a href={`https://teams.microsoft.com/l/chat/0/0?users=${user.Email}`}>
                          <FontIcon aria-label="Compass" iconName="TeamsLogo16" className={teamsIconClass} />
                        </a>
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