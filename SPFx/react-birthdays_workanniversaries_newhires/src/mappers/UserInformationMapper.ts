import { UserProfileInformation } from '../constants';
import { UserInformation } from '../models';

export class UserInformationMapper
{
    public static mapToUserInformations(usersSharePoint: any[]): UserInformation[]
    {
        debugger;
        const mappedUsers = usersSharePoint.map(userSharePoint => 
            new UserInformation({
                title: userSharePoint.Title,
                jobTitle: userSharePoint.JobTitle,
                userTitle: userSharePoint.User.Title,
                userEmail: userSharePoint.User.EMail,
                birthDate: new Date(),
                hireDate: new Date(),
            }));

        return mappedUsers;
    }
}