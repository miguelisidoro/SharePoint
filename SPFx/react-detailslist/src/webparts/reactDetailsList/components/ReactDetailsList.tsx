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
      showPanel: false
    };

    this._columns = [
      { key: 'column1', name: 'Name', fieldName: 'Name', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Email', fieldName: 'Email', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column3', name: 'MobileNumber', fieldName: 'MobileNumber', minWidth: 100, maxWidth: 200, isResizable: true },
    ];
  }

  public async componentDidMount(): Promise<void> {
    // Populate with items for demos.

    await this.loadContacts();
  }

  private async loadContacts() {
    console.log("componentDidMount: begin...");
    this._allItems = [];

    console.log("componentDidMount: before getting contacts...");

    let contacts: IContact[] = await this.sharePointServiceProvider.getContacts();

    console.log("componentDidMount: after getting contacts...");

    this._allItems = contacts;

    console.log("All items count: " + this._allItems.length);

    this.setState({ items: this._allItems, selectionDetails: this._getSelectionDetails() });
  }

  // On New Item
  private onNewItem(e: React.MouseEvent<HTMLElement>) {
    e.preventDefault();
    this.setState({
      panelMode: panelMode.New,
      showPanel: true,
      readOnly: false
    });
  }
  // On Delete
  private async onDeleteItem() {
    // try {
    //   this.setState({
    //     isDeleteting: true
    //   });
    //   await this.spService.deleteImmobilizedRequest(Number(this.state.selectedItem.id));
    //   var partialImobilizedRequests = this.state.partialImobilizedRequests.filter(x => x.id !== this.state.selectedItem.id);
    //   //TODO check if its deleted by getting service with id
    //   //clears selection
    //   this._selection.selectToKey(null, true);
    //   this.setState({
    //     partialImobilizedRequests: partialImobilizedRequests,
    //     hasErrorOnDelete: false,
    //     errorMessage: "",
    //     isDeleteting: false,
    //     showConfirmDelete: false,
    //     hasError: false,
    //     selectedItem: null,
    //     disableCommandOption: true
    //   });

    // } catch (error) {
    //   Log.error(LOG_SOURCE, error, this.context.serviceScope);
    //   console.log("Error on _onDeletePedidoCompra,", error.message);
    //   this.setState({
    //     hasErrorOnDelete: true,
    //     errorMessage: `${error.message}`,
    //     isDeleteting: false
    //   });
    // }
  }

  // On Edit item
  private onEditItem(e: React.MouseEvent<HTMLElement>) {
    e.preventDefault();
    this.setState({
      panelMode: panelMode.Edit,
      showPanel: true,
      readOnly: false

    });
  }
  private onViewItem(e: React.MouseEvent<HTMLElement>) {
    e.preventDefault();
    this.setState({
      panelMode: panelMode.Edit,
      showPanel: true,
      readOnly: true,
    });
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

  private refreshData(ev: React.MouseEvent<HTMLElement>) {
    this.loadContacts();
    //clears selection
    this._selection.selectToKey(null, true);
    this.setState({
      selectedItem: null,
    });
  }

  public render(): JSX.Element {

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
                addItemToList={this.refreshData}
              />
            </div>
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
    alert(`Item invoked: ${item.Name}`);
  };
}
