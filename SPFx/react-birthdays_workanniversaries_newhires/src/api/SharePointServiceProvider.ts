import { Log } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SharePointFieldNames } from "../constants";
import { UserInformationMapper } from "../mappers";
import { UserInformation } from "../models";
import * as moment from 'moment';
import { InformationType } from "../enums";

const LOG_SOURCE: string = "BirthdaysWorkAnniverariesNewHires";
const PROFILE_IMAGE_URL: string =
    "/_layouts/15/userphoto.aspx?size=M&accountname=";
const DEFAULT_PERSONA_IMG_HASH: string = "7ad602295f8386b7615b582d87bcc294";
const DEFAULT_IMAGE_PLACEHOLDER_HASH: string =
    "4a48f26592f4e1498d7a478a4c48609c";
const MD5_MODULE_ID: string = "8494e7d7-6b99-47b2-a741-59873e42f16f";

export class SharePointServiceProvider {
    private _sharePointRelativeListUrl: string;
    private _numberOfItemsToShow: number;
    private _numberOfDaysToRetrieve: number;

    constructor(private _context: WebPartContext,
        sharePointRelativeListUrl: string,
        numberOfItemsToShow: number,
        numberOfDaysToRetrieve: number) {

        // Setup Context to PnP JS
        sp.setup({
            spfxContext: this._context
        });

        this._sharePointRelativeListUrl = sharePointRelativeListUrl;
        this._numberOfItemsToShow = numberOfItemsToShow;
        this._numberOfDaysToRetrieve = numberOfDaysToRetrieve;

        this.onInit();
    }

    private async onInit() { }

    // Sort birthdays by birthdate
    private SortUsersByBirthDate(users: UserInformation[]) {
        return users.sort((a, b) => {
            if (a.birthDate > b.birthDate) {
                return 1;
            }
            if (a.birthDate < b.birthDate) {
                return -1;
            }
            return 0;
        });
    }

    private SortUsersByHireDate(users: UserInformation[]) {
        return users.sort((a, b) => {
            if (a.hireDate > b.hireDate) {
                return 1;
            }
            if (a.hireDate < b.hireDate) {
                return -1;
            }
            return 0;
        });
    }

    private SortUsers(anniversaries: UserInformation[], informationType: InformationType): UserInformation[] {
        if (informationType === InformationType.Birthdays) {
            anniversaries = this.SortUsersByBirthDate(anniversaries);
        }
        else {
            anniversaries = this.SortUsersByHireDate(anniversaries);
        }

        return anniversaries;
    }

    // Get users anniversaries or new hires
    // Important NOTE: All dates are stored with year 2000
    public async getAnniversariesOrHireDates(informationType: InformationType): Promise<UserInformation[]> {
        let currentDate: Date, today: string, currentMonth: string, currentDay: number;
        let filter: string;
        let otherYearUsers: UserInformation[], currentYearUsers: UserInformation[], allUsers: UserInformation[];

        try {
            const filterField: string = informationType === InformationType.Birthdays ?
                SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;

            today = '2000-' + moment().format('MM-DD');
            currentMonth = moment().format('MM');
            currentDate = moment(today).toDate();
            currentDay = parseInt(moment(today).format('DD'));
            let currentDatewithDaysToRetrieve = currentDate;
            if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                // For anniversaries, add the number of days to the current date (we want to show future dates)
                currentDatewithDaysToRetrieve.setDate(currentDate.getDate() + this._numberOfDaysToRetrieve);
            }
            else {
                // For new hires, decrease the number of days from the current date (we want to show past dates)
                currentDatewithDaysToRetrieve.setDate(currentDate.getDate() - this._numberOfDaysToRetrieve);
            }
            let currentDatewithDaysToRetrieveYear = moment(currentDatewithDaysToRetrieve).format('YYYY');

            if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                let filterEndDate;
                // If end date is from next year, get anniversaries from both years
                if (currentDatewithDaysToRetrieveYear === '2001') {
                    // filter end date is calculated, taking one year from the currentDatewithDaysToRetrieve since all dates are stores in year 2000
                    filterEndDate = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD');

                    filter = `(${filterField} ge '${today}' and ${filterField} le '2000-12-31') or (${filterField} ge ` +
                        `'2000-01-01' and ${filterField} le '${filterEndDate}')`;
                }
                else {
                    // simpler filter => just filter dates greater than today and less or equal to the end date (today + number of days to retrieve)
                    filterEndDate = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD');
                    filter = `${filterField} ge '${today}' and ${filterField} le '${filterEndDate}'`;
                }
            }
            else // New Hires
            {
                let filterStartDate;
                // If end date is from next year, get new hires from both years
                if (currentDatewithDaysToRetrieveYear === '1999') {
                    // filter end date is calculated, adding one year to the currentDatewithDaysToRetrieve since all dates are stores in year 2000
                    filterStartDate = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD');

                    filter = `(${filterField} le '${today}' and ${filterField} ge '2000-01-01') or (${filterField} le ` +
                        `'2000-12-31' and ${filterField} ge '${filterStartDate}')`;
                }
                else {
                    // simpler filter => just filter dates less or equal to today and greater than start date (today - number of days to retrieve)
                    filterStartDate = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD');
                    filter = `${filterField} le '${today}' and ${filterField} ge '${filterStartDate}'`;
                }
            }

            const usersSharePoint = await sp.web
                .getList(this._sharePointRelativeListUrl)
                .items.select(
                    SharePointFieldNames.Id,
                    SharePointFieldNames.Title,
                    SharePointFieldNames.UserTitle,
                    SharePointFieldNames.UserEmail,
                    SharePointFieldNames.JobTitle,
                    SharePointFieldNames.BirthDate,
                    SharePointFieldNames.HireDate)
                .expand(SharePointFieldNames.User)
                .filter(filter)
                .top(this._numberOfItemsToShow)
                .usingCaching()
                .get();

            if (usersSharePoint && usersSharePoint.length > 0) {

                allUsers = UserInformationMapper.mapToUserInformations(usersSharePoint);

                if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                    // If end date is from next year, first, select all anniversaries from current year to sort
                    // Then, select all anniversaries of the remaining months from next year
                    // Finally, contact both arrays and return
                    if (currentDatewithDaysToRetrieveYear === '2001') {
                        // get the anniversaries from the current year (months are >= currentMonth)
                        currentYearUsers = allUsers.filter(b => moment(b.birthDate).month() + 1 >= parseInt(currentMonth));
                        currentYearUsers = this.SortUsers(currentYearUsers, informationType);
                        otherYearUsers = allUsers.filter(b => moment(b.birthDate).month() + 1 < parseInt(currentMonth));
                        otherYearUsers = this.SortUsers(otherYearUsers, informationType);
                        // Join the 2 arrays
                        allUsers = currentYearUsers.concat(otherYearUsers);
                    }
                    else {
                        allUsers = this.SortUsers(allUsers, informationType);
                    }
                }
                else
                {
                    // New Hires
                    // If end date is from previous year, first, select all new hires from current year to sort
                    // Then, select all of the remaining months from previous year
                    // Finally, contact both arrays, starting by previous year users and return
                    if (currentDatewithDaysToRetrieveYear === '1999') {
                        // get the new hires from the previous year (months are >= currentMonth)
                        currentYearUsers = allUsers.filter(b => moment(b.birthDate).month() + 1 <= parseInt(currentMonth));
                        currentYearUsers = this.SortUsers(currentYearUsers, informationType);
                        otherYearUsers = allUsers.filter(b => moment(b.birthDate).month() + 1 > parseInt(currentMonth));
                        otherYearUsers = this.SortUsers(otherYearUsers, informationType);
                        // Join the 2 arrays
                        allUsers = currentYearUsers.concat(otherYearUsers);
                    }
                    else {
                        allUsers = this.SortUsers(allUsers, informationType);
                    }
                }
            }

            return allUsers;
        } catch (error) {
            Log.error(LOG_SOURCE, error, this._context.serviceScope);
            throw new Error(error.message);
        }
    }
}