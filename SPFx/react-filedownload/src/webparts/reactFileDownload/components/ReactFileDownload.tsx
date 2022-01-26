import * as React from 'react';
import { IReactFileDownloadProps } from './IReactFileDownloadProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {
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
  Panel,
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
  TextField,
  PanelType,
  IDropdownOption,
  Dropdown
} from '@fluentui/react';
import { DropdownMenuItemType } from 'office-ui-fabric-react';
import { IReactFileDownloadState } from './IReactFileDownloadState';

const yearOptions = [
  { key: '2022', text: '2022', itemType: DropdownMenuItemType.Header },
  { key: '2021', text: '2021' },
  { key: '2020', text: '2020' },
];

const monthInputItems = [
  '1',
  '2',
  '3',
  '4',
  '5',
  '6',
  '7',
  '8',
  '9',
  '10',
  '11',
  '12',
]
export default class ReactFileDownload extends React.Component<IReactFileDownloadProps, IReactFileDownloadState> {

  constructor(props) {
    super(props);

    this.onYearChange = this.onYearChange.bind(this);
  }

  private onYearChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    event.preventDefault();
    this.setState({year: item.key.toString()})
  };

  private onMonthChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    event.preventDefault();
    this.setState({month: item.key.toString()})
  };

  public render(): React.ReactElement<IReactFileDownloadProps> {
    return (
      <Stack>
        <Dropdown options={yearOptions} placeholder="Ano" onChange={this.onYearChange} />
        <Dropdown options={yearOptions} placeholder="Mês" onChange={this.onMonthChange} />
        <TextField
          label='Senha'
          onChange={this.onPasswordChange}
        />
      </Stack>
    );
  }
}
