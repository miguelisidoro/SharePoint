import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IContact } from '../models/IContact'
import { IContactSharePoint } from "../models/IContactSharePoint";
import { ContactMapper } from "../mappers/ContactMapper";
import { SharePointFieldNames, SharePointListNames } from "../constants";

export default class SharePointServiceProvider {
  constructor(private _context: WebPartContext) {
    // Setuo Context to PnPjs and MSGraph
    sp.setup({
      spfxContext: this._context
    });

    this.onInit();
  }

  private async onInit() { }

  public async getContacts(): Promise<IContact[]> {
    try {
      const results: IContactSharePoint[] = await sp.web.lists
        .getByTitle(SharePointListNames.Contacts)
        .items.select(SharePointFieldNames.Title, SharePointFieldNames.Email, SharePointFieldNames.MobileNumber)
        .usingCaching()
        .orderBy(SharePointFieldNames.Title)
        .get();

      let contacts: IContact[];

      contacts = ContactMapper.MapToContacts(results);

      return contacts;
    } catch (error) {
      throw new Error(error.message);
    }
  }

  public async getContactDetailById(id: number): Promise<IContact> {
    try {
      const results: IContactSharePoint = await sp.web.lists
        .getByTitle(SharePointListNames.Contacts)
        .items.getById(id)
        .get();

      let contact: IContact;

      contact = ContactMapper.MapToContact(results);

      return contact;
    } catch (error) {
      throw new Error(error.message);
    }
  }

  // Add Contact to SharePoint List
  public addContact = async (contact: IContact): Promise<void> => {
    try {
      var t = null;
      await sp.web.lists.getByTitle(SharePointListNames.Contacts).items.add({
        Title: `${contact.Name}`,
        Email: `${contact.Email}`,
        Telemovel: `${contact.MobileNumber}`
      });
    } catch (error) {
      throw new Error(error.message);
    }
  }

  public updateContact = async (contact: IContact): Promise<void> => {
    try {
      await sp.web.lists
        .getByTitle(SharePointListNames.Contacts)
        .items.getById(Number(contact.Id))
        .update({
          Title: `${contact.Name}`,
          Email: `${contact.Email}`,
          Telemovel: `${contact.MobileNumber}`
        });
    } catch (error) {
      throw new Error(error.message);
    }
  }
}


