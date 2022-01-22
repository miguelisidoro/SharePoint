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
  TextField
} from '@fluentui/react';
import { IReactDetailsItemPanelState } from './IReactDetailsItemPanelState';
import { IContact, panelMode } from '../../../models';
import * as strings from 'ReactDetailsListWebPartStrings';
import SharePointServiceProvider from '../../../api/SharePointServiceProvider';
import { IReactDetailsItemPanelProps } from './IReactDetailsItemPanelProps';

export default class ReactDetailsItemPanel extends React.Component<IReactDetailsItemPanelProps, IReactDetailsItemPanelState> {

  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props) {
    super(props);

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context);
  }

  public async componentDidMount(): Promise<void> {
    if (this.props.mode === panelMode.Edit) {
      this.getContactDetails();
    }
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <Separator></Separator>
        <PrimaryButton text={this.props.mode === panelMode.Edit ? strings.PrimaryButtonLabelSave : strings.PrimaryButtonLabelInsert} disabled={this.props.readOnly} onClick={this.onSave} style={{ marginRight: "8px" }}>
          {/* {this.state.isAddingPedidoCompra ? <Spinner size={SpinnerSize.small} /> : 'Criar'} */}
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
    //TODO: implement get contact details
    const contact: IContact = await this.sharePointServiceProvider.getContactDetailById(
      this.props.Contact.Id
    );

    if (contact) {
      this.setState ({Contact: contact});
    }
  }

  private validateForm(): boolean {
    let contact: IContact = this.state.Contact;
    if (contact == null)
      throw "Erro ao processar o pedido";
    if (contact.Email == null)
      throw "Invalid Email";
    if (contact.MobileNumber == null)
      throw "Invalid Mobile Number";
    if (contact.Name == null)
      throw "Invalid Name";

    return true;
  }

  private async onSave(ev: React.MouseEvent<HTMLButtonElement>) {
    ev.preventDefault();
    try {
      if (this.validateForm()) {
        switch (this.props.mode) {
          // add immobilized request
          case (panelMode.New):
            try {
              await this.sharePointServiceProvider.addContact(this.state.Contact);
              this.props.addItemToList(ev, this.state.Contact);
              this.setState({ showPanel: false });
              // this.setState({
              //   showPanelConfirmation: true,
              // });
            } catch (error) {
              this.setState({ errorMessage: error });
            }
            break;
          //edit immobilized request
          case (panelMode.Edit):
            try {
              await this.sharePointServiceProvider.updateContact(this.state.Contact);
              this.setState({ showPanel: false });
            } catch (error) {
              this.setState({ errorMessage: error });
            }
            break;
          default:
            break;
        }
      }
      //TODO ERRROR
      else {

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
        Name: newValue.substring(0, 40)
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
        Email: newValue.substring(0, 40)
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
        MobileNumber: newValue.substring(0, 40)
      }
    });
  };

  // Render
  public render(): React.ReactElement<IReactDetailsItemPanelProps> {
    const requiredTag = <label style={{ color: "rgb(168, 0, 0)" }}>*</label>;
    return (
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
    );
  }
}