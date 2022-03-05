import * as moment from 'moment';
import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsWebPartStrings';
import { PersonaInformation, UserInformation } from '@app/models';
import { InformationType } from '@app/enums';
import { DateHelper } from '@app/helpers';
import { UserProfileInformation } from '@app/constants';

export class PersonaInformationMapper {
    public static mapToPersonaInformations(users: UserInformation[], informationType: InformationType): PersonaInformation[] {
        const mappedPersonas = users.map(user => {
            let secondaryText: string;

            if (informationType === InformationType.Birthdays)
            {
                secondaryText = DateHelper.isDateToday(user.BirthDate) ? strings.TodayLabel : DateHelper.formatDateToString(user.BirthDate);
            }
            else if (informationType === InformationType.WorkAnniversaries)
            {
                secondaryText = DateHelper.isDateToday(user.WorkAnniversary) ? strings.TodayLabel : DateHelper.formatDateToString(user.WorkAnniversary);
            }
            else //New Collaborators
            {
                secondaryText = DateHelper.isDateToday(user.HireDate) ? strings.TodayLabel : DateHelper.formatDateToString(user.HireDate);
            }

            return new PersonaInformation({
                imageUrl: UserProfileInformation.profilePictureUrlPrefix + user.Email,
                text: user.Title,
                secondaryText: secondaryText,
                userPrincipalName: user.Email,
            });
        });

        return mappedPersonas;
    }
}