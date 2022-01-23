import { IContact, IContactSharePoint, panelMode } from "../../../models";

export interface IReactDetailsListState {
    //items: IReactDetailsListItem[];
    items: IContact[];
    showPanel: boolean;
    readOnly: boolean;
    isDeleteting: boolean;
    panelMode: panelMode;
    selectionDetails: string;
    showConfirmDelete: boolean;
    disableCommandSelectionOption: boolean;
    selectedItem: IContact;
  }