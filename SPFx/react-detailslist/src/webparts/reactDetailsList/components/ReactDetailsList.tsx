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

export interface IReactDetailsListState {
  items: IReactDetailsListItem[];
  selectionDetails: string;
}

export class ReactDetailsList extends React.Component<IReactDetailsListProps, IReactDetailsListState> {
  private _selection: Selection;
  private _allItems: IReactDetailsListItem[];
  private _columns: IColumn[];

  constructor(props: IReactDetailsListProps) {
    super(props);

    sp.setup({
      spfxContext: this.context
    });

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    this._allItems = [];

    this.state = {
      items: this._allItems,
      selectionDetails: this._getSelectionDetails(),
    };

    this._columns = [
      { key: 'column1', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Value', fieldName: 'value', minWidth: 100, maxWidth: 200, isResizable: true },
    ];
  }

  public async componentDidMount(): Promise<void> {
    // Populate with items for demos.
    this._allItems = [];
    for (let i = 0; i < 200; i++) {
      this._allItems.push({
        key: i,
        name: 'Item ' + i,
        value: i,
      });
    }

    this.setState({items: this._allItems, selectionDetails: this._getSelectionDetails()});
  }

  public render(): JSX.Element {

    if (this.state.items != null && this.state.items.length > 0) {
      const { items, selectionDetails } = this.state;

      return (
        <div>
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
        </div>
      );
    }
    else
    {
      return (<div></div>);
    }
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IReactDetailsListItem).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems,
    });
  };

  private _onItemInvoked = (item: IReactDetailsListItem): void => {
    alert(`Item invoked: ${item.name}`);
  };
}
