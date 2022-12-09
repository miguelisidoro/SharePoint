import { UserProfileInformation } from '../constants';
import { PersonInformation } from '../data';
import { Microsoft365GroupUser } from '../data/Microsoft365GroupUser';

export class PersonInformationMapper
{
    public static mapToPersonInformations(users: Microsoft365GroupUser[]): PersonInformation[]
    {
        const mappedUsers = users.map(user => 
            new PersonInformation({
                imageUrl: UserProfileInformation.profilePictureUrlPrefix + user.userPrincipalName,
                text: user.name,
                secondaryText: user.jobTitle,
                userPrincipalName: user.userPrincipalName
            }));

        return mappedUsers;
    }
}