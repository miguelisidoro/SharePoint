import { MSGraphClient } from "@microsoft/sp-http"
import { WebPartContext } from "@microsoft/sp-webpart-base"
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Microsoft365Group } from "../data";
import { VirtualizedComboBox } from "@fluentui/react";
import { Microsoft365GroupMapper } from "../mapper/Microsoft365GroupMapper";
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { Providers } from '@microsoft/mgt-element';

import { graph } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";

export class GraphServiceProvider {

    constructor(private _context: WebPartContext) {
        //Providers.globalProvider = new SharePointProvider(_context);
        graph.setup({
            spfxContext: _context
        });
    }

    public async getCurrentUserGroups(): Promise<Microsoft365Group[]> {
        debugger;
        const members = await graph.groups.get();
        
        return null;

        //let currentUserGroups: Microsoft365Group[];

        //const currentUserGroupsGraph = await Providers.globalProvider.graph.api('/me/transitiveMemberOf').select('').get();

        //eturn currentUserGroups;
    }
}