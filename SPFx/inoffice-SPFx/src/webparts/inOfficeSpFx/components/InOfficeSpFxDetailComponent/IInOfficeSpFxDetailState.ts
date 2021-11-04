import { IColumn, IContextualMenuProps } from "office-ui-fabric-react";

import {IInOfficeAppointment} from "../../../../interfaces/IInOfficeAppointment" 

//import { IListViewItems } from "./IListViewItems";
import { panelMode } from "../../../../spservices/IEnumPanel";

export interface IInOfficeSpFxState {
  items:  IInOfficeAppointment[];
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
  selectItem: IInOfficeAppointment;
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
