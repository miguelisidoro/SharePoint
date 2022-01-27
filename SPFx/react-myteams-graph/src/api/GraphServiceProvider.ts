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

    public async getMicrosoft365Groups(): Promise<Microsoft365Group[]> {
        debugger;
        const allGroups: MicrosoftGraph.Group[] = await graph.groups
        .filter("groupTypes/any(a:a%20eq%20'unified')")
        .select("displayName, id")
        .usingCaching()
        .get();
        
        let allMappedGroups: Microsoft365Group[] = Microsoft365GroupMapper.MapToMicrosoft365Groups(allGroups);

        //const currentUserGroupsGraph = await Providers.globalProvider.graph.api('/me/transitiveMemberOf').select('').get();

        return allMappedGroups;
    }
}