import { UserProfileInformation } from '../constants';
import { PersonaInformation, UserInformation } from '../models';

export class PersonaInformationMapper
{
    public static mapToPersonaInformations(users: UserInformation[]): PersonaInformation[]
    {
        const mappedPersonas = users.map(user => 
            new PersonaInformation({
                imageUrl: UserProfileInformation.profilePictureUrlPrefix + user.userEmail,
                text: user.userTitle,
                secondaryText: user.jobTitle,
                userPrincipalName: user.userEmail
            }));

        return mappedPersonas;
    }
}