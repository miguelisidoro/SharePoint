import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {IContact} from '../models/IContact'

export default class spservices {
    constructor(private _context: WebPartContext) {
        // Setuo Context to PnPjs and MSGraph
        sp.setup({
            spfxContext: this._context
        });

        this.onInit();
    }

    private async onInit() { }

    public async getContacts(): Promise<IContact[]> {
        const results: IContact[] = await sp.web.lists
          .getByTitle("Contacts")
          .items.select("Title", "Email", "Telemovel")
          .usingCaching()
          .orderBy("DescricaoItem")
          .get();

        return results;
    }
}


