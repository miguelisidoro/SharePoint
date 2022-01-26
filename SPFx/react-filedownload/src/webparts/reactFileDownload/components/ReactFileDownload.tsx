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
  { key: '2022', text: '2022' },
  { key: '2021', text: '2021' },
  { key: '2020', text: '2020' },
];

const monthOptions = [
  { key: '1', text: '1' },
  { key: '2', text: '2' },
  { key: '3', text: '3' },
  { key: '4', text: '4' },
  { key: '5', text: '5' },
  { key: '6', text: '6' },
  { key: '7', text: '7' },
  { key: '8', text: '8' },
  { key: '9', text: '9' },
  { key: '10', text: '10' },
  { key: '11', text: '11' },
  { key: '12', text: '12' },
];
export default class ReactFileDownload extends React.Component<IReactFileDownloadProps, IReactFileDownloadState> {

  constructor(props) {
    super(props);

    this.onYearChange = this.onYearChange.bind(this);
    this.onMonthChange = this.onMonthChange.bind(this);
    this.onPasswordChange = this.onPasswordChange.bind(this);
  }

  private onYearChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    event.preventDefault();
    this.setState({year: item.key.toString()});
  };

  private onMonthChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    event.preventDefault();
    this.setState({month: item.key.toString()});
  };

  private onPasswordChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {event.preventDefault();
    this.setState({password: newValue});
  };

  private async onDownloadReceipt(ev: React.MouseEvent<HTMLButtonElement>) {
    ev.preventDefault();
    console.log("onDownloadReceipt");
  }

  public render(): React.ReactElement<IReactFileDownloadProps> {
    return (
      <Stack>
        <Dropdown options={yearOptions} placeholder="Ano" onChange={this.onYearChange} />
        <Dropdown options={monthOptions} placeholder="Mês" onChange={this.onMonthChange} />
        <TextField
          label='Senha'
          onChange={this.onPasswordChange}
        />
        <PrimaryButton text='Visualizar recibo' onClick={this.onDownloadReceipt} style={{ marginRight: "8px" }}>
        </PrimaryButton>
      </Stack>
    );
  }
}
