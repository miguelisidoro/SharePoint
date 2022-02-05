import { UserProfileInformation } from '../constants';
import { UserInformation } from '../models';

export class UserInformationMapper
{
    public static mapToUserInformations(usersSharePoint: any[]): UserInformation[]
    {
        const mappedUsers = usersSharePoint.map(userSharePoint => 
            new UserInformation({
                Title: userSharePoint.Title,
                JobTitle: userSharePoint.JobTitle,
                Email: userSharePoint.EMail,
                BirthDate: userSharePoint.BirthDate,
                HireDate: userSharePoint.HireDate,
            }));

        return mappedUsers;
    }
}