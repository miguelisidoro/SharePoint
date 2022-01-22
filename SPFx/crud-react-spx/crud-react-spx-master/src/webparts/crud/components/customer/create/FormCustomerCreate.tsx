import * as React from 'react';
import { IFormCustomerCreateState } from './IFormCustomerCreateState';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  TextField,
  DefaultButton,
  MessageBar,
  MessageBarType,
  MessageBarButton
} from 'office-ui-fabric-react';
import { CustomersDataProvider } from '../sharePointDataProvider/CustomersDataProvider';
import { Customer } from '../Models/Customer';
import { ICustomer } from '../Models/ICustomer';
export default class FormCustomerCreate extends React.Component<{}, IFormCustomerCreateState> {
  private _customersDataProvider:CustomersDataProvider;
  private _customer:Customer;
  constructor(props){
    super(props);
    this._customersDataProvider=new CustomersDataProvider({});
    this.state = {
      isBusy: false,
      customer:  new Customer(),
      customersDataProvider: this._customersDataProvider,
      messageSended: false,
      _goBack:props.state._goBack,
      _reload:props.state._reload,
    };
  }

  public render(): React.ReactElement<{}> {
    return (
      <div>
              <TextField disabled={this.state.isBusy} label="Customer Name"  name="text" value={this.state.customer.name} onChange={this._onChange} />
              <div >
                <DefaultButton disabled={this.state.isBusy } onClick={this._CreateCustomer}>Save Customer</DefaultButton>
                <br></br>
               <DefaultButton  onClick={this.state._goBack} >Cancel</DefaultButton>
           
              </div>
      </div>
    );
  }

  private _onChange = (event: React.ChangeEvent<HTMLInputElement>) : void => {
    
    let value= event.target.value;
    const {customer}=   this.state;
    customer.name=value;
    this.setState({customer:customer});
  }

  private _CreateCustomer = async (event: React.MouseEvent<HTMLButtonElement, MouseEvent>) : Promise<void> => {
   
    const {customer}=   this.state;
    this.setState({isBusy:true});
    this._customersDataProvider.createItem(customer).then((customers: ICustomer[]) => {
       this.state._reload();
    });
    
    this.state._goBack();
   
}
}

