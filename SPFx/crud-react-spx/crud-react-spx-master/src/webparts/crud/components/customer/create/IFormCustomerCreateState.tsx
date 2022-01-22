import { ICustomer } from "../Models/ICustomer";
import { ICustomersDataProvider } from "../sharePointDataProvider/ICustomersDataProvider";

export interface IFormCustomerCreateState {
  isBusy: boolean;
  customer: ICustomer;
  messageSended: boolean;
  customersDataProvider:ICustomersDataProvider;
  _goBack:VoidFunction;
  _reload:VoidFunction;
}
