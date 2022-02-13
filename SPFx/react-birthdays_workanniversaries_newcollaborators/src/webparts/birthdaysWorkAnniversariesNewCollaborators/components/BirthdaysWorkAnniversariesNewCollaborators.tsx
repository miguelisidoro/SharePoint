import * as React from 'react';
import styles from './BirthdaysWorkAnniversariesNewCollaborators.module.scss';
import { IBirthdaysWorkAnniversariesNewCollaboratorsProps } from './IBirthdaysWorkAnniversariesNewCollaboratorsProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SharePointServiceProvider } from '../../../api';
import { UserInformation } from '../../../models';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsWebPartStrings';

import {
  CommandBar,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IContextualMenuProps,
  IconButton,
  ImageFit,
  Label,
  Link,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  PrimaryButton,
  SearchBox,
  Separator,
  ShimmeredDetailsList,
  Spinner,
  SpinnerSize,
  Stack,
  getTheme,
  Dropdown,
  IDropdownStyles,
  IDropdownOption,
  themeRulesStandardCreator,
  IPersonaProps
} from '@fluentui/react';
import { LivePersona } from "@pnp/spfx-controls-react/lib/controls/LivePersona";

import { IBirthdaysWorkAnniverariesNewCollaboratorsState } from './IBirthdaysWorkAnniverariesNewCollaboratorsState';
import { ServiceScope } from '@microsoft/sp-core-library';
import { PersonaInformationMapper } from '../../../mappers/PersonaInformationMapper';
import { InformationDisplayType, InformationType } from '../../../enums';

const personaProps: IPersonaProps = {
  size: PersonaSize.size48,
  styles: {
    root: {
      width: 325,
      margin: 5,
    },
  },
};
export default class BirthdaysWorkAnniversariesNewCollaborators extends React.Component<IBirthdaysWorkAnniversariesNewCollaboratorsProps, IBirthdaysWorkAnniverariesNewCollaboratorsState> {
  
  private sharePointServiceProvider: SharePointServiceProvider;
  private _serviceScope: ServiceScope;

  constructor(props: IBirthdaysWorkAnniversariesNewCollaboratorsProps) {
    super(props);

    this._serviceScope = null;

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context,
      this.props.sharePointRelativeListUrl,
      this.props.numberOfItemsToShow,
      this.props.numberOfDaysToRetrieve);

    this.state = {
      users: null,
    };

    this.loadUsers = this.loadUsers.bind(this);
    this.isWebPartConfigured = this.isWebPartConfigured.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadUsers();
  }

  /// Checkes if web part is properly configured
  private isWebPartConfigured(): boolean
  {
    const isWebPartConfigured = (this.props.sharePointRelativeListUrl != null && this.props.sharePointRelativeListUrl != undefined
      && this.props.showMoreUrl != null && this.props.showMoreUrl != undefined
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

      // users = await this.sharePointServiceProvider.getAnniversariesOrNewCollaborators(informationType, InformationDisplayType.TopResults);

      users = await this.sharePointServiceProvider.getAnniversariesOrNewCollaborators(informationType, InformationDisplayType.TopResults);

      if (users != null && users.length > 0) {
        let usersPersonInformation = PersonaInformationMapper.mapToPersonaInformations(users, informationType);

        this.setState({
          users: usersPersonInformation,
        });
      }
    }
  }
  
  public render(): React.ReactElement<IBirthdaysWorkAnniversariesNewCollaboratorsProps> {
    if (this.isWebPartConfigured()) {
      if (this.state.users !== null && this.state.users.length > 0) {
        return (
          <div>
            <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateTitleProperty} />
            {
              this.state.users !== null && this.state.users.map(user =>
                <LivePersona serviceScope={this._serviceScope} upn={user.userPrincipalName}
                  template={
                    <>
                      <Persona {...user} {...personaProps} />
                    </>
                  }
                  context={this.props.context}
                />
              )
            }
            <Link href={this.props.showMoreUrl}>
              {strings.ShowMoreLabel}
            </Link>
          </div>
        );
      }
      else {
        let noUsersMessage;
        const informationType: InformationType = InformationType[this.props.informationType];
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
