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
}
