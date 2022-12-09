import * as React from 'react';

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
  PanelType
} from '@fluentui/react';
import { IReactDetailsItemPanelState } from './IReactDetailsItemPanelState';
import { IContact, Contact, panelMode } from '../../../models';
import * as strings from 'ReactDetailsListWebPartStrings';
import SharePointServiceProvider from '../../../api/SharePointServiceProvider';
import { IReactDetailsItemPanelProps } from './IReactDetailsItemPanelProps';
import * as tsStyles from "./ReactDetailsListStyles";

export default class ReactDetailsItemPanel extends React.Component<IReactDetailsItemPanelProps, IReactDetailsItemPanelState> {
  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props) {
    super(props);

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context);

    this.state = ({
      showPanel: true,
      readOnly: true,
      visible: true,
      multiline: true,
      primaryButtonLabel: strings.PrimaryButtonLabelSave,
      disableButton: false,
      errorMessage: '',
      Contact: null,
      showPanelConfirmation: false,
      isLoading: true,
    });

    this.onCancel = this.onCancel.bind(this);
    this.onSave = this.onSave.bind(this);
    this.getContactDetails = this.getContactDetails.bind(this);
    this.initializeEmptyContact = this.initializeEmptyContact.bind(this);
    this.validateForm = this.validateForm.bind(this);
    this._onChangeEmail = this._onChangeEmail.bind(this);
    this._onChangeMobileNumber = this._onChangeMobileNumber.bind(this);
    this._onChangeName = this._onChangeName.bind(this);
    this._onDismiss = this._onDismiss.bind(this);
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    console.log("Panel componentDidMount")
    if (this.props.mode === panelMode.Edit) {
      this.getContactDetails();
    }
    else {
      this.initializeEmptyContact();
    }
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <Separator></Separator>
        <PrimaryButton text={this.props.mode === panelMode.Edit ? strings.PrimaryButtonLabelSave : strings.PrimaryButtonLabelInsert} disabled={this.props.readOnly} onClick={this.onSave} style={{ marginRight: "8px" }}>
        </PrimaryButton>
        <DefaultButton text={strings.PrimaryButtonLabelCancel} onClick={this.onCancel} />
      </div>
    );
  }

  // Cancel Panel
  private onCancel(ev: React.MouseEvent<HTMLButtonElement>) {
    ev.preventDefault();
    this.props.onDismiss();
  }

  private async getContactDetails(): Promise<void> {
    const contact: IContact = await this.sharePointServiceProvider.getContactDetailById(
      this.props.Contact.Id
    );

    if (contact) {
      this.setState({ Contact: contact, isLoading: false });
    }
  }

  private async initializeEmptyContact(): Promise<void> {
    const contact: IContact = new Contact(
      {
        Name: '',
        Email: '',
        MobileNumber: ''
      }
    );

    this.setState({ Contact: contact, isLoading: false });
  }

  private validateForm(): boolean {
    let contact: IContact = this.state.Contact;
    if (contact === null)
      throw "Error processing request!";
    if (contact.Name === null || contact.Name === '')
      throw "Invalid Name!";
    if (contact.Email === null || contact.Email === '')
      throw "Invalid Email!";
    if (contact.MobileNumber === null || contact.MobileNumber === '')
      throw "Invalid Mobile Number!";

    return true;
  }

  private async onSave(ev: React.MouseEvent<HTMLButtonElement>) {
    ev.preventDefault();
    console.log("onSave");
    try {
      if (this.validateForm()) {
        switch (this.props.mode) {
          // add contact
          case (panelMode.New):
            try {
              await this.sharePointServiceProvider.addContact(this.state.Contact);
              this.props.refreshData(ev, this.state.Contact);
              this.props.onDismiss();
            } catch (error) {
              this.setState({ errorMessage: error });
            }
            break;
          //edit contact
          case (panelMode.Edit):
            try {
              await this.sharePointServiceProvider.updateContact(this.state.Contact);
              this.props.refreshData(ev, this.state.Contact);
              this.props.onDismiss();
            } catch (error) {
              this.setState({ errorMessage: error });
            }
            break;
          default:
            break;
        }
      }
    } catch (e) {
      this.setState({
        errorMessage: e
      });
    }
  }

  private _onChangeName = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    event.preventDefault();
    this.setState({
      Contact: {
        ...this.state.Contact,
        Name: newValue
      }
    });
  };

  private _onChangeEmail = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    event.preventDefault();
    this.setState({
      Contact: {
        ...this.state.Contact,
        Email: newValue
      }
    });
  };

  private _onChangeMobileNumber = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    event.preventDefault();
    this.setState({
      Contact: {
        ...this.state.Contact,
        MobileNumber: newValue
      }
    });
  };

  private _onDismiss = (ev?: React.SyntheticEvent<HTMLElement, Event>) => {
    ev.preventDefault();
    ev.stopPropagation();
    this.props.onDismiss();
  };

  // Render
  public render(): React.ReactElement<IReactDetailsItemPanelProps> {
    console.log("render panel");
    if (this.state.Contact != null) {
      return (
        <Panel
          closeButtonAriaLabel="Close"
          isOpen={this.state.showPanel}
          type={PanelType.medium}
          onDismiss={this.props.onDismiss}
          isFooterAtBottom={true}
          headerText={this.props.mode == panelMode.Edit ? (this.props.readOnly ? strings.PanelHeaderTextVisualize : strings.PanelHeaderTextEdit) : strings.PanelHeaderTextAdd}
          onRenderFooterContent={this._onRenderFooterContent}
        >
          {
            this.state.errorMessage && (
              <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>
            )
          }
          {
            this.state.isLoading ? (
              <Stack disableShrink styles={tsStyles.stackStyles}   >
                <Stack.Item align="stretch">
                  <Spinner size={SpinnerSize.large} />
                </Stack.Item>
              </Stack>
            ) : (
              <Stack>
                <TextField
                  label={strings.NameFieldLabel}
                  readOnly={this.props.readOnly}
                  value={this.state.Contact.Name}
                  onChange={this._onChangeName}
                />
                <TextField
                  label={strings.EmailFieldLabel}
                  readOnly={this.props.readOnly}
                  value={this.state.Contact.Email}
                  onChange={this._onChangeEmail}
                />
                <TextField
                  label={strings.MobileNumberFieldLabel}
                  readOnly={this.props.readOnly}
                  value={this.state.Contact.MobileNumber}
                  onChange={this._onChangeMobileNumber}
                />
              </Stack>
            )
          }
        </Panel>
      );
    }
    else {
      return (<div></div>);
    }
  }
}