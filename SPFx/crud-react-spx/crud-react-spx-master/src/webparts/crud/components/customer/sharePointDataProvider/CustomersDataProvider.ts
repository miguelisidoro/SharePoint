import { IWebPartContext } from '@microsoft/sp-webpart-base';
import {ICustomer} from '../Models/ICustomer';
import {ICustomersDataProvider} from './ICustomersDataProvider';
import { sp, ItemAddResult } from "@pnp/sp";
const BASEURL="https://pedrommm.sharepoint.com/sites/BlogNet/";
const LIST_CUSTOMER="Customer";
export  class CustomersDataProvider implements ICustomersDataProvider {

  constructor(props: {}) {

    sp.setup({
      sp: {
        headers: {
          Accept: "application/json;odata=verbose",
        },
        baseUrl: BASEURL
      },
    });
    
  }
    private _listCustomersUrl: string;
    private _listsUrl: string;
    private _webPartContext: IWebPartContext;
    private _customers: ICustomer[];
    public set webPartContext(value: IWebPartContext) {
      this._webPartContext = value;
      this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;
    }
  
    public get webPartContext(): IWebPartContext {
      return this._webPartContext;
    }
    public getItems(): Promise<ICustomer[]> {
      let customers: ICustomer[]=[];
      // get all the customers from the list customers in SharePoint
      return sp.web.lists.getByTitle(LIST_CUSTOMER).items.get().then((result_customers: any[]) => {
        result_customers.forEach(customer => {
       if(typeof customer!='undefined' && customer)
       {
         //uncommented and the next statement validates all the fields to not allow nulls values
        //if(typeof customer.Title!='undefined' && customer.Title 
       // && typeof customer.Id!='undefined' && customer.Id &&
       // typeof customer.LastName!='undefined' && customer.LastName){
          customers.push({name:replaceNullsByEmptyString(customer.Title),key:customer.Id,value:replaceNullsByEmptyString(customer.LastName)});
      // }
        
       }
      
     });
      return customers;
    });
   
    }
    public createItem(itemCreated: ICustomer): Promise<ICustomer[]> {
      let customers: ICustomer[]=[];
      // add an item to the list
      return  sp.web.lists.getByTitle(LIST_CUSTOMER).items.add({
          Title:  itemCreated.name,
          LastName: itemCreated.value
        }).then((iar: ItemAddResult) => {
          console.log(iar);
          customers.push(itemCreated);
          return customers;
        
        });
    }
    public updateItem(itemUpdated: ICustomer): Promise<ICustomer[]> {
       // update an item to the list
      let customers: ICustomer[]=[];
     let id=  itemUpdated.key;
     return sp.web.lists.getByTitle(LIST_CUSTOMER).items.getById(id).update({
      Title: itemUpdated.name,
      LastName:itemUpdated.value
  }).then((result_customers) => {
    console.log(result_customers);
     customers.push(itemUpdated);
    return customers;
  });
    }
    public deleteItem(itemDeleted: ICustomer): Promise<ICustomer[]> {
      throw new Error("Method not implemented.");
    }
   
  }
  function replaceNullsByEmptyString( value) {
    return (value == null) ? "" : value
}