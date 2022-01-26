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
import { DropdownMenuItemType, IDropdownStyles } from 'office-ui-fabric-react';
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

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 100 },
  dropdownOptionText: { overflow: 'visible', whiteSpace: 'normal' },
  dropdownItem: { height: 'auto' },
};

const textFieldStyes = () => {
  return {
    root: {
      maxWidth: '300px'
    }
  }
};

export default class ReactFileDownload extends React.Component<IReactFileDownloadProps, IReactFileDownloadState> {

  constructor(props) {
    super(props);

    this.onYearChange = this.onYearChange.bind(this);
    this.onMonthChange = this.onMonthChange.bind(this);
    this.onPasswordChange = this.onPasswordChange.bind(this);
    this.onDownloadReceipt = this.onDownloadReceipt.bind(this);
    this.base64ToArrayBuffer = this.base64ToArrayBuffer.bind(this);
    this.createAndDownloadBlobFile = this.createAndDownloadBlobFile.bind(this);
  }

  private onYearChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    event.preventDefault();
    this.setState({ year: item.key.toString() });
  };

  private onMonthChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    event.preventDefault();
    this.setState({ month: item.key.toString() });
  };

  private onPasswordChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    event.preventDefault();
    this.setState({ password: newValue });
  };

  private onDownloadReceipt(ev: React.MouseEvent<HTMLButtonElement>) {
    ev.preventDefault();
    console.log("onDownloadReceipt");

    const base64Pdf = "";

    const arrayBuffer = this.base64ToArrayBuffer(base64Pdf);
    this.createAndDownloadBlobFile(arrayBuffer, 'testName');
  }

private base64ToArrayBuffer(base64: string) {
  const binaryString = window.atob(base64); // Comment this if not using base64
  const bytes = new Uint8Array(binaryString.length);
  return bytes.map((byte, i) => binaryString.charCodeAt(i));
}

private createAndDownloadBlobFile(body, filename, extension = 'pdf') {
  const blob = new Blob([body]);
  const fileName = `${filename}.${extension}`;
  const link = document.createElement('a');
  // Browsers that support HTML5 download attribute
  if (link.download !== undefined) {
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}

  public render(): React.ReactElement < IReactFileDownloadProps > {
  return(
      <Stack>
        <Label>Ano</Label>
        <Dropdown styles={dropdownStyles} options={yearOptions} placeholder="Ano" onChange={this.onYearChange} />
        <Label>Mês</Label>
        <Dropdown styles={dropdownStyles} options={monthOptions} placeholder="Mês" onChange={this.onMonthChange} />
        <TextField
          label='Senha'
          type='password'
          onChange={this.onPasswordChange}
          styles = {textFieldStyes}
        />
        <Separator></Separator>
        <PrimaryButton text='Visualizar recibo' onClick={this.onDownloadReceipt} style={{ width:"200px"}}>
        </PrimaryButton>
      </Stack >
    );
  }
}
