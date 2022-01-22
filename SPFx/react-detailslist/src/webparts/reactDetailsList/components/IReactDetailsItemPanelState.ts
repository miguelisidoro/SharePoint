import { IContact } from "../../../models";

export interface IReactDetailsItemPanelState {
    showPanel: boolean;
    readOnly: boolean;
    visible: boolean;
    multiline: boolean;
    errorMessage: string;
    primaryButtonLabel: string;
    disableButton: boolean;
    Contact: IContact;
    //ImmobilizedListsIds: IImmobilizedListsIds;
    showPanelConfirmation: boolean;
    isLoading:boolean;
  }