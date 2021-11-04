import { IColumn, IContextualMenuProps } from "office-ui-fabric-react";

import { IListViewItems } from "./IListViewItems";
import { panelMode } from "../../../../spservices/IEnumPanel";

export interface IInOfficeSpFxState {
  items:  IListViewItems[];
  isLoading:boolean;
  disableCommandOption:boolean;
  disableCommandEdit:boolean;
  disableCommandFlow:boolean;
  disableCommandDelete:boolean;
  showConfirmDelete:boolean;
  showPanelAdd:boolean;
  showPanelEdit:boolean;
  showPanelView:boolean;
  showPanelDetalhe:boolean;
  showPanelFlow: boolean;
  selectItem: IListViewItems;
  panelMode: panelMode;
  hasError: boolean;
  errorMessage: string;
  hasMore: boolean;
  columns: IColumn[];
  title: string;
  hideDialogDelete:boolean;
  hasErrorOnDelete:boolean;
  isDeleteting:boolean;
  isConfirmingFlow :boolean;

}
