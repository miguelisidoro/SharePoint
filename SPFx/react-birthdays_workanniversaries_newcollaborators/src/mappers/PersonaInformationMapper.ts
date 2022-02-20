import * as moment from 'moment';
import { UserProfileInformation } from '../constants';
import { PersonaInformation, UserInformation } from '../models';
import { InformationType } from '../enums';

import * as strings from 'BirthdaysWorkAnniversariesNewCollaboratorsWebPartStrings';

export class PersonaInformationMapper {
    // Returns if a date is today
    private static isDateToday(inputDate: Date): boolean {
        const currentDay = moment().date();
        const currentMonth = moment().month() + 1;
        const inputDateDay = moment(inputDate).date();
        const inputDateMonth = moment(inputDate).month() + 1;

        const isDateToday = (currentDay === inputDateDay && currentMonth === inputDateMonth) ? true : false;

        return isDateToday;
    }

    private static formatDateToString(inputDate: Date): string {
        return moment(inputDate, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('Do MMMM');
    }

    public static mapToPersonaInformations(users: UserInformation[], informationType: InformationType): PersonaInformation[] {
        const mappedPersonas = users.map(user => {
            let secondaryText: string;

            if (informationType === InformationType.Birthdays)
            {
                secondaryText = this.isDateToday(user.BirthDate) ? strings.TodayLabel : this.formatDateToString(user.BirthDate);
            }
            else if (informationType === InformationType.WorkAnniversaries)
            {
                secondaryText = this.isDateToday(user.WorkAnniversary) ? strings.TodayLabel : this.formatDateToString(user.WorkAnniversary);
            }
            else //New Collaborators
            {
                secondaryText = this.isDateToday(user.HireDate) ? strings.TodayLabel : this.formatDateToString(user.HireDate);
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