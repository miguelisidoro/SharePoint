import * as React from 'react';
import styles from './InOfficeSpFx.module.scss';
import { IInOfficeSpFxProps } from './IInOfficeSpFxProps';
import { IInOfficeSpFxState } from "./IInOfficeSpFxState";
import { escape } from '@microsoft/sp-lodash-subset';
import { panelMode } from "../../../spservices/IEnumPanel";

const LOG_SOURCE: string = "In Office SPFx";
const IMAGE_LIST_NO_ITEMS: string =
  "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/images/emptyfolder/empty_list.svg";

import {
  CommandBar,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dialog,
  DialogFooter,
  DialogType,
  FontIcon,
  IColumn,
  IContextualMenuProps,
  Icon,
  IconButton,
  IconType,
  Label,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  PrimaryButton,
  SearchBox,
  Selection,
  SelectionMode,
  Separator,
  ShimmeredDetailsList,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
  getTheme,
  mergeStyles
} from "office-ui-fabric-react";
import { FontSizes, FontWeights } from "@uifabric/styling";
import InfiniteScroll from "react-infinite-scroller";
import * as strings from "InOfficeSpFxWebPartStrings";
import * as tsStyles from "./InOfficeSpFxStyles";

export default class InOfficeSpFx extends React.Component<
IInOfficeSpFxProps,
IInOfficeSpFxState> {

  private _spservices: _spservices = new _spservices(this.props.context);
  private _pagedResults: PagedItemCollection<any[]>;
  private _selection: Selection;
  private _itemIdParameter: string = "";
  private _disableForm: boolean =this.props.pnanelMode == panelMode.View ? true : false;
  private theme = getTheme();
  private _isScrolling:boolean = false;
  private _isSorting:boolean = false;
  private _isSearching:boolean = false;

  private _selection: Selection;

  constructor(props: IInOfficeSpFxProps) {
    super(props);

    const columns: IColumn[] = [
      {
        key: "column1",
        name: "File_x0020_Type",
        className: tsStyles.classNames.fileIconCell,
        iconClassName: tsStyles.classNames.fileIconHeaderIcon,
        iconName: "TextDocument",
        isIconOnly: true,
        fieldName: "name",
        minWidth: 20,
        maxWidth: 20,
        onColumnClick: this._onColumnClick,
        onRender: (item: IListViewItems) => {
          const renderFileType: JSX.Element = (
            <FontIcon iconName="TextDocument" className={styles.iconClass} />
          );
          return renderFileType;
        }
      },
      {
        name: strings.IDFieldLabel,
        key: "column2",
        fieldName: "Id",
        minWidth: 20,
        maxWidth: 40,
        isResizable: true,
        isSorted: true,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        name: "",
        key: "column3",
        minWidth: 40,
        maxWidth: 40,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        onRender: (item: IListViewItems) => {
          return (
            <div className={tsStyles.classNames.centerColumn}>
              <IconButton
                className={styles.buttonMenu}
                checked={true}
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
      {
        key: "column4",
        name: strings.DataFieldLabel,
        fieldName: "DataPedido",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        key: "column5",
        name: strings.SolicitanteFieldLabel,
        fieldName: "Solicitante",
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true,
        onRender: (item: IListViewItems) => {
          const userProfileInfo = {
            imageUrl: item.SolicitantePhotoUrl,
            text: item.SolicitanteDisplayName
          };
          const renderFileType: JSX.Element = (
            <Persona {...userProfileInfo} size={PersonaSize.size24} />
          );
          return renderFileType;
        }
      },
      {
        key: "column6",
        name: strings.FornecedorFieldLabel,
        fieldName: "Fornecedor",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        key: "column7",
        name: strings.DescricaoFornecedorFieldLabel,
        fieldName: "DescricaoFornecedor",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        key: "column8",
        name: "Total",
        fieldName: "Total",
        minWidth: 90,
        maxWidth: 90,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        //  onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        key: "column9",
        name: strings.EstadoFieldLabel,
        fieldName: "EstadoPedido",
        minWidth: 100,
        maxWidth: 110,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true,
        onRender: this._renderEstadoPedido
      },
      {
        key: "column10",
        name: strings.NumeroPedidoCompraLabel,
        fieldName: "Numero",
        minWidth: 100,
        maxWidth: 110,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true,
        onRender: (item: IListViewItems) => {
          let renderNumeroPedidoSAP: JSX.Element = null;
          if (item.EstadoPedido === estadoPedido.Liberado) {
            renderNumeroPedidoSAP = (
              <Label className={styles.label}>{item.Numero}</Label>
            );
          }
          return renderNumeroPedidoSAP;
        }
      },
      {
        key: "column11",
        name: strings.EmpresaFieldLabel,
        fieldName: "Empresa",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      },
      {
        key: "column12",
        name: strings.DescricaoFieldLabel,
        fieldName: "DescricaoEmpresa",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true
      }
    ];

    this.state = {
      isLoading: false,
      items: [],
      disableCommandOption: true,
      disableCommandEdit: true,
      disableCommandDelete: true,
      disableCommandFlow: true,
      errorMessage: "",
      hasError: false,
      panelMode: panelMode.New,
      selectItem: undefined,
      showPanelAdd: false,
      showPanelEdit: false,
      showPanelView: false,
      showPanelDetalhe: false,
      showPanelFlow: false,
      hasMore: false,
      columns: columns,
      title: this.props.title,
      showConfirmDelete: false,
      hideDialogDelete: true,
      hasErrorOnDelete: false,
      isDeleteting: false,
      isConfirmingFlow: false
    };
  
    // handler selection Item List
    this._selection = new Selection({
      onSelectionChanged: () => {
        this._getSelectionDetails();
      }
    });
  }

/**
   * Determines whether column click on
   */
 private _onColumnClick = async (
  ev: React.MouseEvent<HTMLElement>,
  column: IColumn
): Promise<void> => {
  // tslint:disable-next-line: no-shadowed-variable
  if (this._isSorting) return ;
  this._isSorting = true;
  const { columns, hasMore } = this.state;
  let { items } = this.state;
  let newItems: IListViewItems[] = [];
  const newColumns: IColumn[] = columns.slice();
  const currColumn: IColumn = newColumns.filter(
    currCol => column.key === currCol.key
  )[0];
  newColumns.forEach((newCol: IColumn) => {
    if (newCol === currColumn) {
      currColumn.isSortedDescending = !currColumn.isSortedDescending;
      currColumn.isSorted = true;
    } else {
      newCol.isSorted = false;
      newCol.isSortedDescending = true;
    }
  });
  if (hasMore) {
    // has more  items to load get items sorted by clicked columns and direction
    // the pnpjs REST API the sort parameter is controled by Ascending (true or false)
    //  is diferent from Column.isSortedDescending indication on DataListView Columns Properties
    // for that is  !currColumn.isSortedDescending
    await this._getPedidoCompraDetalhe(
      currColumn.fieldName,
      !currColumn.isSortedDescending
    );
    this._isSorting = false;
  } else {
    items = this.state.items;
    // Sort Items
    newItems = this._copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColumns,
      items: newItems
    });
    this._isSorting = false;
  }
};
/**
 * Copys and sort
 * @template T
 * @param items
 * @param columnKey
 * @param [isSortedDescending]
 * @returns and sort
 */
private _copyAndSort<T>(
  items: T[],
  columnKey: string,
  isSortedDescending?: boolean
): T[] {
  const key = columnKey as keyof T;
  return items
    .slice(0)
    .sort((a: T, b: T) =>
      (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
    );
}
/**
 * Gets selection details
 */
private _getSelectionDetails() {
  const selectionCount = this._selection.getSelectedCount();
  const selectedItem = this._selection.getSelection()[0] as IListViewItems;

  switch (selectionCount) {
    case 0:
      this.setState({
        selectItem: null,
        disableCommandOption: true,
        disableCommandDelete: true,
        disableCommandEdit: true
      });
      break;
    case 1:
      this.setState({
        selectItem: this._selection.getSelection()[0] as IListViewItems,
        disableCommandOption: false,
        disableCommandDelete: this._disableForm ? true : false,
        disableCommandEdit: this._disableForm ? true : false
      });

      break;
    default:
  }
}
  
  public render(): React.ReactElement<IInOfficeSpFxProps> {
    return (
      // <div className={ styles.inOfficeSpFx }>
      //   <div className={ styles.container }>
      //     <div className={ styles.row }>
      //       <div className={ styles.column }>
      //         <span className={ styles.title }>Welcome to SharePoint!</span>
      //         <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
      //         <p className={ styles.description }>{escape(this.props.description)}</p>
      //         <a href="https://aka.ms/spfx" className={ styles.button }>
      //           <span className={ styles.label }>Learn more</span>
      //         </a>
      //       </div>
      //     </div>
      //   </div>
      // </div>

      <div>
          <InfiniteScroll
            pageStart={0}
            threshold={50}
            //loadMore={this._getPedidosCompraNextPage}
            //hasMore={this.state.hasMore}
            useWindow={false}//funciona
          >
            <ShimmeredDetailsList
              items={this.state.items}
              compact={false}
              columns={this.state.columns}
              selectionMode={SelectionMode.single}
              setKey="items"
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              enableShimmer={this.state.isLoading}
              enableUpdateAnimations={false}
              selection={this._selection}
            />
          </InfiniteScroll>
          {this.state.items.length == 0 && !this.state.isLoading && (
            <Stack tokens={{ childrenGap: 0 }}>
              <Stack.Item align="center">
                <Icon
                  className={styles.noListItemsImageStyle}
                  iconType={IconType.Image}
                  imageProps={{
                    src: IMAGE_LIST_NO_ITEMS
                  }}
                />
              </Stack.Item>
              <Stack.Item align="center">
                <Label className={styles.title}>
                  {strings.NoItemsListMessage}
                </Label>
              </Stack.Item>
            </Stack>
          )}
        </div>
    );
  }
}
