import * as React from 'react';
import styles from './BirthdaysWorkAnniverariesNewHires.module.scss';
import { IBirthdaysWorkAnniverariesNewHiresProps } from './IBirthdaysWorkAnniverariesNewHiresProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SharePointServiceProvider } from '../../../api';
import { UserInformation } from '../../../models';

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

import { IBirthdaysWorkAnniverariesNewHiresState } from './IBirthdaysWorkAnniverariesNewHiresState';
import { ServiceScope } from '@microsoft/sp-core-library';
import { PersonaInformationMapper } from '../../../mappers/PersonaInformationMapper';

const personaProps: IPersonaProps = {
  size: PersonaSize.size48,
  styles: {
    root: {
      width: 325,
      margin: 5,
    },
  },
};

export default class BirthdaysWorkAnniverariesNewHires extends React.Component<IBirthdaysWorkAnniverariesNewHiresProps, IBirthdaysWorkAnniverariesNewHiresState> {

  private sharePointServiceProvider: SharePointServiceProvider;
  private _serviceScope: ServiceScope;

  constructor(props: IBirthdaysWorkAnniverariesNewHiresProps) {
    super(props);

    this._serviceScope = null;

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context,
      this.props.sharePointRelativeListUrl,
      this.props.numberOfItemsToShow,
      this.props.numberOfDaysToRetrieve);

    this.state = {
      users: null,
    };
  }

  public async componentDidMount(): Promise<void> {
    // Populate with items for demos.

    await this.loadUsers();
  }

  private async loadUsers() {
    debugger;
    if (this.props.sharePointRelativeListUrl != null && this.props.sharePointRelativeListUrl != undefined) {
      let users: UserInformation[] = await this.sharePointServiceProvider.getUserBirthDays();

      const usersPersonInformation = PersonaInformationMapper.mapToPersonaInformations(users);

      this.setState({
        users: usersPersonInformation,
      });
    }
  }

  public render(): React.ReactElement<IBirthdaysWorkAnniverariesNewHiresProps> {
    if (this.state.users !== null) {
      return (
        <div>
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
        </div>
      );
    }
    else {
      return (<div></div>);
    }
  }
}
