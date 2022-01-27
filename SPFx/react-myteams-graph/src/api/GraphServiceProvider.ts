import { MSGraphClient } from "@microsoft/sp-http"
import { WebPartContext } from "@microsoft/sp-webpart-base"
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Microsoft365Group } from "../data";
import { VirtualizedComboBox } from "@fluentui/react";

export class GraphServiceProvider {

    constructor(private _context: WebPartContext) {
    }

    public async getCurrentUserGroups(): Promise<Microsoft365Group[]> {

        let groups: Microsoft365Group[];

        // this._context.msGraphClientFactory
        //     .getClient()
        //     .then((client: MSGraphClient): void => {
        //         // get information about the current user from the Microsoft Graph
        //         client
        //             .api('/me/transitiveMemberOf/')
        //             //.top(5)
        //             .orderby("displayName asc")
        //             .get((error, groups: MicrosoftGraph.Group[], rawResponse?: any) => {

        //             });

        return groups;
    }
}