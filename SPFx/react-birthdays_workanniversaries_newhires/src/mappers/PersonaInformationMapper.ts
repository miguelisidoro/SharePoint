import * as moment from 'moment';
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
                secondaryText: moment(user.birthDate, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('Do MMMM'),
                userPrincipalName: user.userEmail
            }));

        return mappedPersonas;
    }
}