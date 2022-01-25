import * as React from 'react';
import { Announced } from '@fluentui/react/lib/Announced';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyles } from '@fluentui/react/lib/Styling';
import { Text } from '@fluentui/react/lib/Text';
import { IReactDetailsListProps } from './IReactDetailsListProps'
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IContact, panelMode } from "../../../models";
import SharePointServiceProvider from '../../../api/SharePointServiceProvider';
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
import * as strings from 'ReactDetailsListWebPartStrings';
import { IReactDetailsListState } from './IReactDetailsListState';
import ReactDetailsItemPanel from './ReactDetailsItemPanel';
import * as tsStyles from "./ReactDetailsListStyles";

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

export interface IReactDetailsListItem {
  key: number;
  name: string;
  value: number;
}

export class ReactDetailsList extends React.Component<IReactDetailsListProps, IReactDetailsListState> {
  private _selection: Selection;
  //private _allItems: IReactDetailsListItem[];
  private _allItems: IContact[];
  private _columns: IColumn[];
  private theme = getTheme();
  private sharePointServiceProvider: SharePointServiceProvider;

  constructor(props: IReactDetailsListProps) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({ selectionDetails: this._getSelectionDetails() });
        this._onSelectionChanged();
      }
    });

    this.sharePointServiceProvider = new SharePointServiceProvider(this.props.context);

    this._allItems = [];

    this.state = {
      items: this._allItems,
      selectionDetails: this._getSelectionDetails(),
      isDeleteting: false,
      panelMode: panelMode.New,
      readOnly: false,
      selectedItem: null,
      disableCommandSelectionOption: true,
      showConfirmDelete: false,
      showPanel: false,
      hasError: false,
      hasErrorOnDelete: false,
      errorMessage: ''
    };

    const listItemMenuProps: IContextualMenuProps = {
      items: [
        {
          key: "0",
          text: strings.CommandbarEditLabel,
          iconProps: { iconName: "Edit" },
          onClick: this.onEditItem.bind(this),
          disabled: this.state ? this.state.disableCommandSelectionOption : true,
        },
        {
          key: "1",
          text: strings.CommandbarViewLabel,
          iconProps: { iconName: "View" },
          onClick: this.onViewItem.bind(this),
        },
        {
          key: "2",
          text: strings.CommandbarDeleteLabel,
          iconProps: { iconName: "Delete" },
          onClick: this._openDialogDelete.bind(this),
          disabled: this.state ? this.state.disableCommandSelectionOption : true,
        }
      ]
    };

    this._columns = [
      { key: 'idColumn', name: 'Id', fieldName: 'Id', minWidth: 10, maxWidth: 200, isResizable: true },
      {
        name: "",
        key: "menuPropsColumn",
        minWidth: 40,
        maxWidth: 40,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        onRender: (item: IContact) => {
          return (
            <div className={tsStyles.classNames.centerColumn}>
              <IconButton
                style={{ backgroundColor: "#1fe0" }}
                iconProps={{ iconName: "MoreVertical" }}
                text={""}
                width="30"
                split={false}
                onMenuClick={this._onListItemIdMenuClick}
                menuIconProps={{ iconName: "" }}
                menuProps={listItemMenuProps}
              />
            </div>
          );
        }
      },
      { key: 'nameColumn', name: 'Name', fieldName: 'Name', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'emailColumn', name: 'Email', fieldName: 'Email', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'mobileNumberColumn', name: 'MobileNumber', fieldName: 'MobileNumber', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    this.onNewItem = this.onNewItem.bind(this);
    this.onEditItem = this.onEditItem.bind(this);
    this.onViewItem = this.onViewItem.bind(this);
    this.onDeleteItem = this.onDeleteItem.bind(this);
    this.onDismissPanel = this.onDismissPanel.bind(this);
    this.loadContacts = this.loadContacts.bind(this);
    this._openDialogDelete = this._openDialogDelete.bind(this);
    this._onListItemIdMenuClick = this._onListItemIdMenuClick.bind(this);
    this._closeDialogDelete = this._closeDialogDelete.bind(this);
  }

  public async componentDidMount(): Promise<void> {
    // Populate with items for demos.

    await this.loadContacts();
  }

  private _onListItemIdMenuClick = (ev: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, button: any) => {
    button.menuProps.items[0].disabled = this.state.disableCommandSelectionOption;
    button.menuProps.items[2].disabled = this.state.disableCommandSelectionOption;
  }

  private async loadContacts() {
    console.log("componentDidMount: begin...");
    this._allItems = [];

    console.log("componentDidMount: before getting contacts...");

    let contacts: IContact[] = await this.sharePointServiceProvider.getContacts();

    console.log("componentDidMount: after getting contacts...");

    this._allItems = contacts;

    console.log("All items count: " + this._allItems.length);

    this.setState({ 
      items: this._allItems, 
      selectionDetails: this._getSelectionDetails()
     });
  }

  // On New Item
  private onNewItem(e: React.MouseEvent<HTMLElement>) {
    e.preventDefault();
    console.log("New Item");
    
    this.setState({
      panelMode: panelMode.New,
      showPanel: true,
      readOnly: false
    });
    console.log("New Item State Set");
  }

  private _closeDialogDelete = async () => {
    this.setState({ showConfirmDelete: false });
  }

  // On Delete
  private async onDeleteItem() {
    try {
      this.setState({
        isDeleteting: true
      });
      await this.sharePointServiceProvider.deleteContact(this.state.selectedItem.Id);
      var contactsAterDelete = this.state.items.filter(x => x.Id !== this.state.selectedItem.Id);
      //clears selection
      this._selection.selectToKey(null, true);
      this.setState({
        items: contactsAterDelete,
        hasErrorOnDelete: false,
        errorMessage: '',
        isDeleteting: false,
        showConfirmDelete: false,
        hasError: false,
        selectedItem: null,
        disableCommandSelectionOption: true,
      });

    } catch (error) {
      console.log("Error on onDeleteItem,", error.message);
      this.setState({
        hasErrorOnDelete: true,
        errorMessage: `${error.message}`,
        isDeleteting: false
      });
    }
  }

  // On Edit item
  private onEditItem(e: React.MouseEvent<HTMLElement>) {
    e.preventDefault();
    console.log("Edit Item");
    this.setState({
      panelMode: panelMode.Edit,
      showPanel: true,
      readOnly: false
    });
    console.log("Edit Item State Set");
  }
  private onViewItem(e: React.MouseEvent<HTMLElement>) {
    e.preventDefault();
    console.log("View Item");
    this.setState({
      panelMode: panelMode.Edit,
      showPanel: true,
      readOnly: true,
    });
    console.log("View Item State Set");
  }

  private _openDialogDelete = async () => {
    this.setState({ showConfirmDelete: true });
  }

  private async onDismissPanel(ev?: React.SyntheticEvent<HTMLElement>) {
    if (ev)
      ev.preventDefault();
    this.setState({
      showPanel: false
    });
  }

  refreshData = async (ev: React.MouseEvent<HTMLElement>) => {
    console.log("refreshData");
    await this.loadContacts();
    //clears selection
    this._selection.selectToKey(null, true);
    this.setState({
      selectedItem: null,
    });
  }

  public render(): JSX.Element {

    console.log("render");
    if (this.state.items != null && this.state.items.length > 0) {
      const { items, selectionDetails } = this.state;

      return (
        <div>
          <CommandBar
            items={[
              {
                key: 'newItem',
                name: strings.CommandbarNewLabel,
                iconProps: {
                  iconName: 'Add',
                },
                onClick: this.onNewItem,
              },
              {
                key: 'edit',
                name: strings.CommandbarEditLabel,
                iconProps: {
                  iconName: 'Edit'
                },
                onClick: this.onEditItem,
                disabled: this.state.disableCommandSelectionOption,
              },
              {
                key: 'view',
                name: strings.CommandbarViewLabel,
                iconProps: {
                  iconName: 'View'
                },
                onClick: this.onViewItem,
                disabled: this.state.disableCommandSelectionOption,
              },
              {
                key: 'delete',
                name: strings.CommandbarDeleteLabel,
                iconProps: {
                  iconName: 'Delete'
                },
                onClick: this._openDialogDelete,
                disabled: this.state.disableCommandSelectionOption,
              }
            ]}
          />

          <div className={exampleChildClass}>{selectionDetails}</div>
          <Text>
            Note: While focusing a row, pressing enter or double clicking will execute onItemInvoked, which in this
            example will show an alert.
          </Text>
          <Announced message={selectionDetails} />
          <TextField
            className={exampleChildClass}
            label="Filter by name:"
            onChange={this._onFilter}
            styles={textFieldStyles}
          />
          <Announced message={`Number of items after filter applied: ${items.length}.`} />
          <MarqueeSelection selection={this._selection}>
            <DetailsList
              items={items}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selection={this._selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
              onItemInvoked={this._onItemInvoked}
            />
          </MarqueeSelection>
          {
            this.state.showPanel &&
            <div>
              <ReactDetailsItemPanel
                mode={this.state.panelMode}
                Contact={this.state.selectedItem}
                onDismiss={this.onDismissPanel}
                showPanel={this.state.showPanel}
                context={this.props.context}
                readOnly={this.state.readOnly}
                refreshData={this.refreshData}
              />
            </div>
          }
          {
          this.state.showConfirmDelete && (
            <Dialog
              hidden={!this.state.showConfirmDelete}
              onDismiss={this._closeDialogDelete}
              dialogContentProps={{
                type: DialogType.largeHeader,
                title: strings.DeleteTitle
              }}
              modalProps={{
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
              }}>
              <Stack>
                <Label>{strings.DeleteLabelId}</Label>
                <TextField
                  disabled
                  defaultValue={this.state.selectedItem ? String(this.state.selectedItem.Id) : ""}
                  style={{ color: this.theme.palette.neutralPrimary }}
                />
                <Label>{strings.DeleteLabelName}</Label>
                <TextField
                  disabled
                  defaultValue={this.state.selectedItem ? this.state.selectedItem.Name : ""}
                  style={{ color: this.theme.palette.neutralPrimary }}
                />
                {this.state.isDeleteting && <Spinner size={SpinnerSize.medium} label={strings.DeletingMessage} />}
                {this.state.hasErrorOnDelete && <MessageBar messageBarType={MessageBarType.error}>{this.state.errorMessage}</MessageBar>}
              </Stack>
              <DialogFooter>
                <PrimaryButton disabled={this.state.isDeleteting} onClick={this.onDeleteItem} text={strings.PrimaryButtonLabelDelete} />
                <DefaultButton onClick={this._closeDialogDelete} text={strings.PrimaryButtonLabelCancel} />
              </DialogFooter>
            </Dialog>
          )
        }
        </div>
      );
    }
    else {
      return (<div></div>);
    }
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IContact).Name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  // handles the on selection changed in the DetailsList control
  private _onSelectionChanged() {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        this.setState({
          selectedItem: null,
          disableCommandSelectionOption: true
        });
        break;
      case 1:
        let contact = (this._selection.getSelection()[0] as IContact);
        this.setState({
          selectedItem: this._selection.getSelection()[0] as IContact,
          disableCommandSelectionOption: false
        });
        break;
      default:
        this.setState({
          disableCommandSelectionOption: true
        });
    }
  }

  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this._allItems.filter(i => i.Name.toLowerCase().indexOf(text) > -1) : this._allItems,
    });
  };

  private _onItemInvoked = (item: IContact): void => {
    console.log(`Item invoked: ${item.Name}`);
  };
}
