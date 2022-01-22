import { ICustomer } from "../Models/ICustomer";
import { ICustomersDataProvider } from "../sharePointDataProvider/ICustomersDataProvider";

export interface IFormCustomerEditState {
  isBusy: boolean;
  customer: ICustomer;
  messageSended: boolean;
  customersDataProvider:ICustomersDataProvider;
  showEditCustomerPanel:boolean;
  _goBack:VoidFunction;
}
