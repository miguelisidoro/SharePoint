import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {IContact} from '../models/IContact'
import { IContactSharePoint } from "../models/IContactSharePoint";
import { ContactMapper } from "../mappers/ContactMapper";

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
        const results: IContactSharePoint[] = await sp.web.lists
          .getByTitle("Contacts")
          .items.select("Title", "Email", "Telemovel")
          .usingCaching()
          .orderBy("Title")
          .get();

        let contacts : IContact[];

        contacts = ContactMapper.MapToContact(results);

        return contacts;
    }
}


