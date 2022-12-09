import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Microsoft365GroupUser } from '../data/Microsoft365GroupUser';

export class Microsoft365GroupUserMapper
{
    public static mapToMicrosoft365GroupUsers(users: MicrosoftGraph.User[]): Microsoft365GroupUser[]
    {
        const mappedUsers = users.map(userGraph => 
            new Microsoft365GroupUser({
                email: userGraph.mail,
                name: userGraph.displayName,
                userPrincipalName: userGraph.userPrincipalName,
                jobTitle: userGraph.jobTitle
            }));

        return mappedUsers;
    }
}