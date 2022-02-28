import * as moment from 'moment';
import { UserProfileInformation } from '../constants';
import { PersonaInformation, UserInformation } from '../models';
import { InformationType } from '../enums';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsWebPartStrings';
import { DateHelper } from '../helpers';

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