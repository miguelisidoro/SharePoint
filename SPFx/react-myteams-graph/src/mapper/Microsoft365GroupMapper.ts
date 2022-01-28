import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Microsoft365Group } from '../data';

export class Microsoft365GroupMapper
{
    public static mapToMicrosoft365Groups(groups: MicrosoftGraph.Group[]): Microsoft365Group[]
    {
        const mappedGroups = groups.map(groupGraph => 
            new Microsoft365Group({
                GroupId: groupGraph.id,
                GroupName: groupGraph.displayName,
            }));

        return mappedGroups;
    }
}