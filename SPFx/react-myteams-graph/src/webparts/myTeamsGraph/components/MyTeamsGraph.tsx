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
import { PersonInformation } from '../../../data/PersonInformation'
export default class MyTeamsGraph extends React.Component<IMyTeamsGraphProps, {}> {

  private _userProfileInfo: PersonInformation[];

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
