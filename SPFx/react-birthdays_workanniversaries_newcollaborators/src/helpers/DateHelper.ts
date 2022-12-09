import { InformationType } from "@app/enums";
import { UserInformation } from "@app/models";
import * as moment from "moment";

/// Date helper class
export class DateHelper {
    // Returns if a date is today
    public static isDateToday(inputDate: Date): boolean {
        const currentDay = moment().date();
        const currentMonth = moment().month() + 1;
        const inputDateDay = moment(inputDate).date();
        const inputDateMonth = moment(inputDate).month() + 1;

        const isDateToday = (currentDay === inputDateDay && currentMonth === inputDateMonth) ? true : false;

        return isDateToday;
    }

    public static formatDateToString(inputDate: Date): string {
        return moment(inputDate, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('Do MMMM');
    }

    public static getUserFormattedDate(user: UserInformation, informationType: InformationType, todayDateAsString: string) {
        let formattedDate;
        if (informationType === InformationType.Birthdays) {
            formattedDate = DateHelper.isDateToday(user.BirthDate) ? todayDateAsString : DateHelper.formatDateToString(user.BirthDate);
        } else if (informationType === InformationType.WorkAnniversaries) {
            formattedDate = DateHelper.isDateToday(user.WorkAnniversary) ? todayDateAsString : DateHelper.formatDateToString(user.WorkAnniversary);
        } else { //New Collaborators
            formattedDate = DateHelper.isDateToday(user.HireDate) ? todayDateAsString : DateHelper.formatDateToString(user.HireDate);
        }

        return formattedDate;
    }
}
