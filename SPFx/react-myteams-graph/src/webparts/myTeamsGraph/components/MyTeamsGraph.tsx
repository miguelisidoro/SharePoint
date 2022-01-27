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

const userProfileInfo = {
  imageUrl: '/_layouts/15/userphoto.aspx?size=M&accountname=miguel.isidoro@createdevpt.onmicrosoft.com',
  text: 'Miguel Isidoro'
};

export default class MyTeamsGraph extends React.Component<IMyTeamsGraphProps, {}> {
  
  constructor(props) {
    super(props);
  }
  
  public render(): React.ReactElement<IMyTeamsGraphProps> {
    return (
      <Persona {...userProfileInfo}>
      </Persona>
    );
  }
}
