import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { ICustomer } from "../Models/ICustomer";
export interface IDetailsListCustomersState {
    columns: IColumn[];
    items: ICustomer[];
    selectionDetails: string;
    selectedCustomer: ICustomer;
    showEditCustomerPanel:boolean;
    _goBack:VoidFunction;
    _reloadList?:VoidFunction;
}