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
    private SortBirthdaysByBirthDate(users: UserInformation[]) {
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

    // Get users anniversaries
    public async getAnniversaries(informationType: InformationType): Promise<UserInformation[]> {
        let currentDate: Date, today: string, currentMonth: string, currentDay: number;
        let filter: string;
        let otherMonthsAnniversaries: UserInformation[], currentMonthWithDaysToRetriveAnniversaries: UserInformation[], allAnniversaries: UserInformation[];

        try {

            const filterField: string = informationType === InformationType.Birthdays ? 
            SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;
            
            today = '2000-' + moment().format('MM-DD');
            //today = '2000-12-01';
            currentMonth = moment().format('MM');
            //currentMonth = '12';
            currentDate = moment(today).toDate();
            currentDay = parseInt(moment(today).format('DD'));
            let currentDatewithDaysToRetrieve = currentDate;
            currentDatewithDaysToRetrieve.setDate(currentDate.getDate() + this._numberOfDaysToRetrieve);
            let filterEndDate;
            let endDateYear = moment(currentDatewithDaysToRetrieve).format('YYYY');
            // If end date is from next year, get birthdays from both years
            if (endDateYear === '2001') {
                filterEndDate = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD');
                // filter = "(" + filterField + " ge '" + today + "' and " + filterField + " le '2000-12-31') or (" + filterField + " ge '2000-01-01' and " + filterField + " le '" + filterEndDate + "')";

                filter = `(${filterField} ge '${today}' and ${filterField} le '2000-12-31') or (${filterField} ge ` + 
                `'2000-01-01' and ${filterField} le '${filterEndDate}')`;
            }
            else {
                filterEndDate = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD');
                //filter = "BirthDate ge '" + today + "' and BirthDate le '" + filterEndDate + "'";
                filter = `${filterField} ge '${today}' and ${filterField} le '${filterEndDate}'`;
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

                allAnniversaries = UserInformationMapper.mapToUserInformations(usersSharePoint);

                // If end date is from next year, first, select all birthdays from current year to sort
                // Then, select all birthdays of the remaining months from next year
                // Finally, contact both arrays and return
                if (endDateYear === '2001') {
                    currentMonthWithDaysToRetriveAnniversaries = allAnniversaries.filter(b => moment(b.birthDate).month() + 1 >= parseInt(currentMonth));
                    currentMonthWithDaysToRetriveAnniversaries = this.SortBirthdaysByBirthDate(currentMonthWithDaysToRetriveAnniversaries);
                    otherMonthsAnniversaries = allAnniversaries.filter(b => moment(b.birthDate).month() + 1 < parseInt(currentMonth));
                    otherMonthsAnniversaries = this.SortBirthdaysByBirthDate(otherMonthsAnniversaries);
                    // Join the 2 arrays
                    allAnniversaries = currentMonthWithDaysToRetriveAnniversaries.concat(otherMonthsAnniversaries);
                }
                else {
                     allAnniversaries = this.SortBirthdaysByBirthDate(allAnniversaries);
                }
            }

            return allAnniversaries;
        } catch (error) {
            Log.error(LOG_SOURCE, error, this._context.serviceScope);
            throw new Error(error.message);
        }
    }

    getUserWorkAnniversaries
}