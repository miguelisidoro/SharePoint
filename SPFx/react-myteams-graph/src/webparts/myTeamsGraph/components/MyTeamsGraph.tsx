import * as React from 'react';
import { IMyTeamsGraphProps } from './IMyTeamsGraphProps';
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
  getTheme
} from '@fluentui/react';
import { PersonInformation } from '../../../data/PersonInformation';
import { GraphServiceProvider, SharePointServiceProvider } from '../../../api';
import { Microsoft365Group } from '../../../data';
import { IMyTeamsGraphState } from './IMyTeamsGraphPropsState';
export default class MyTeamsGraph extends React.Component<IMyTeamsGraphProps, IMyTeamsGraphState> {

  private _userProfileInfo: PersonInformation[];
  private sharePointServiceProvider: SharePointServiceProvider;
  private graphServiceProvider: GraphServiceProvider;

  constructor(props) {
    super(props);

    this._userProfileInfo = [{
      imageUrl: '/_layouts/15/userphoto.aspx?size=M&accountname=miguel.isidoro@createdevpt.onmicrosoft.com',
      text: 'Miguel Isidoro'
    },
    {
      imageUrl: '/_layouts/15/userphoto.aspx?size=M&accountname=david.oliveira@createdevpt.onmicrosoft.com',
      text: 'David Oliveira'
    }];

    this.graphServiceProvider = new GraphServiceProvider(this.props.context);

    this.loadCurrentUserGroups = this.loadCurrentUserGroups.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadCurrentUserGroups();
  }

  private async loadCurrentUserGroups() {
    console.log("componentDidMount: begin...");

    let currentUserGroups: Microsoft365Group[] = await this.graphServiceProvider.getCurrentUserGroups();

    this.setState({
      currentUserGroups: currentUserGroups,
    });
  }

  public render(): React.ReactElement<IMyTeamsGraphProps> {
    return (
      <div>
        {
          this._userProfileInfo.map(profile =>
            <Persona {...profile}>
            </Persona>
          )}
      </div>
    );
  }
}
