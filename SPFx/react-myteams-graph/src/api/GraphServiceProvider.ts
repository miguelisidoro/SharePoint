import { MSGraphClient } from "@microsoft/sp-http"
import { WebPartContext } from "@microsoft/sp-webpart-base"
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Microsoft365Group } from "../data";
import { VirtualizedComboBox } from "@fluentui/react";
import { Microsoft365GroupMapper } from "../mapper/Microsoft365GroupMapper";
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { Providers } from '@microsoft/mgt-element';

import { graph, graphGet, GraphQueryable } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";
import { Microsoft365GroupUserMapper } from "../mapper";
import { Microsoft365GroupUser } from "../data/Microsoft365GroupUser";

export class GraphServiceProvider {

    constructor(private _context: WebPartContext) {
        //Providers.globalProvider = new SharePointProvider(_context);
        graph.setup({
            spfxContext: _context
        });
    }

    public async getMicrosoft365Groups(): Promise<Microsoft365Group[]> {
       
        const allGroups: MicrosoftGraph.Group[] = await graph.groups
        .filter("groupTypes/any(a:a%20eq%20'unified')")
        .select("displayName, id")
        .usingCaching()
        .get();
        
        let allMappedGroups: Microsoft365Group[] = Microsoft365GroupMapper.mapToMicrosoft365Groups(allGroups);

        //const currentUserGroupsGraph = await Providers.globalProvider.graph.api('/me/transitiveMemberOf').select('').get();

        return allMappedGroups;
    }

    public async getMicrosoft365GroupMembers(microsoft365GroupId: string): Promise<Microsoft365GroupUser[]> {
        debugger;
        const groupMembers: MicrosoftGraph.User[] = await graph.groups.getById(microsoft365GroupId).members
        .select("displayName, jobTitle, userPrincipalName")
        .usingCaching()
        .get();
        
        let mappedGroupMembers: Microsoft365GroupUser[] = Microsoft365GroupUserMapper.mapToMicrosoft365GroupUsers(groupMembers);

        return mappedGroupMembers;
    }
}