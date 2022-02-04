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
                secondaryText = this.isDateToday(user.birthDate) ? strings.TodayLabel : this.formatDateToString(user.birthDate);
            }
            else
            {
                secondaryText = this.isDateToday(user.hireDate) ? strings.TodayLabel : this.formatDateToString(user.hireDate);
            }

            return new PersonaInformation({
                imageUrl: UserProfileInformation.profilePictureUrlPrefix + user.userEmail,
                text: user.userTitle,
                secondaryText: secondaryText,
                userPrincipalName: user.userEmail,
            })
        });

        return mappedPersonas;
    }
}