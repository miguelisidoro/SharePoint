import * as React from 'react';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import * as lodash from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { CommandBarCustomers } from '../utils/CommandBarCustomers';
import { IDetailsListCustomersState } from './IDetailsListCustomersState';
import { CustomersDataProvider } from '../sharePointDataProvider/CustomersDataProvider';
import {ICustomer} from '../Models/ICustomer';
import FormCustomerEdit from '../edit/FormCustomerEdit';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import Modal from 'office-ui-fabric-react/lib/Modal';
import { DefaultButton, Button } from 'office-ui-fabric-react/lib/Button';
const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px'
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  selectionDetails: {
    marginBottom: '20px'
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};

export class DetailsListCustomers extends React.Component<{}, IDetailsListCustomersState> {
  private _selection: Selection;
  private _allItems: ICustomer[];
  private _customersDataProvider:CustomersDataProvider;
  private _webPartContext: IWebPartContext;
  private showEditCustomerPanel:boolean;
  // Use getId() to ensure that the IDs are unique on the page.
  // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
  private _titleId: string = getId('title');
  private _subtitleId: string = getId('subText');
  constructor(props: {}) {
   
    super(props);
//this is to chage by wev service rest apiget from the list
    this._customersDataProvider=new CustomersDataProvider({});
    this._allItems = this._LoadCustomers();

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Id',
        isIconOnly: false,
        fieldName: 'key',
        minWidth: 30,
        maxWidth: 50,
        data: 'string',
        onColumnClick: this._onColumnClick,
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column2',
        name: 'Last Name',
        fieldName: 'value',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
          showEditCustomerPanel: this.showEditCustomerPanel,
         
        });
      }
    });
  
    this.state = {
      items: this._allItems,
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      showEditCustomerPanel: false,
      selectedCustomer: null,
      _goBack:this._hidePanel
      
      
    };
  }
 
  public render() {
    const { columns, items, selectionDetails,showEditCustomerPanel } = this.state;

    return (
      <Fabric>
        <Separator />
        <CommandBarCustomers  {...this}  />
        <Separator />
        <div className={classNames.controlWrapper}>
        <Stack >
        <TextField label="Filter by name of the customer:" onChange={this._onChangeText} iconProps={{ iconName: 'search' }} styles={controlStyles} />
      </Stack>
         
        </div>
        <div className={classNames.selectionDetails}>{selectionDetails}</div>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
            columns={columns}
            selectionMode={SelectionMode.single}
            getKey={this._getKey}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={(item) => { this._onItemInvoked(item, this); }}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
          />
          
        </MarqueeSelection>
        <div>
        <Panel isOpen={this.state.showEditCustomerPanel} onDismiss={this._hidePanel} type={PanelType.extraLarge} headerText="Edit Customer">
         <FormCustomerEdit {...this}  />
        </Panel>
      </div>
      </Fabric>
    );
  }
  private _LoadCustomers() {
    const items: ICustomer[] = [];
    this._customersDataProvider.getItems().then((customers: ICustomer[]) => {
      customers.forEach(element => {
        items.push({name:element.name,key:element.key,value:element.value});
       });
        return customers;
   
    });
    return items;
  }
//To Update the items in the list
  public componentDidUpdate(previousProps: any, previousState: IDetailsListCustomersState) {
    
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }
  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      
      items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems
    });
  }

  private _onItemInvoked(item: any,value:any): void {
    let itemCustomer=item as ICustomer;
    value.setState( {selectedCustomer:itemCustomer});
    value.setState( {showEditCustomerPanel:true});
   
  }
  

  private _hidePanel = () => {


    const items: ICustomer[] = [];
    this._customersDataProvider.getItems().then((customers: ICustomer[]) => {
      customers.forEach(element => {
        items.push({name:element.name,key:element.key,value:element.value});
       });
       this.setState({ showEditCustomerPanel: false, items:items })
       this.setState({ showEditCustomerPanel: false });
    });
   
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    this.setState( {selectedCustomer:this._selection.getSelection()[0] as ICustomer});
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as ICustomer).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}










