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
        return users.sort( (a, b) => {
          if (a.birthDate > b.birthDate) {
            return 1;
          }
          if (a.birthDate < b.birthDate) {
            return -1;
          }
          return 0;
        });
      }

    // Get users from User Information SharePoint List
    public async getUserBirthDays(): Promise<UserInformation[]> {
        let today: string, currentMonth: string, currentDay: number;
        let filter: string, currentDayWithNumberOfDaysToRetrieve: number, nextYearEndDay: number, nextYearStartDate: string;
        let nextYearEndDate: string;
        let monthsExceptoDecemberBirthdays: UserInformation[], decemberBirthdays: UserInformation[], allBirthDays: UserInformation[];

        try {
            today = '2000-' + moment().format('MM-DD');
            //today = '2000-12-18';
            currentMonth = moment().format('MM');
            //currentMonth = '12';
            currentDay = parseInt(moment().format('DD'));
            filter = "BirthDate ge '" + today + "'";

            // If we are in December, we have to look if there are birthdays in January
            // We have to build a condition to select birthdays from January based on number of days to retrieve
            // We cannot use the year, the year is always 2000
            console.log("currentMonth: " + currentMonth);
            if (currentMonth === '12') {
                currentDayWithNumberOfDaysToRetrieve = currentDay + this._numberOfDaysToRetrieve;
                nextYearStartDate = '2000-01-01';
                //filter = "BirthDate ge '" + today + "' or (BirthDate ge '" + nextYearStartDate + "')";
                if ((currentDayWithNumberOfDaysToRetrieve) > 31) {
                    nextYearStartDate = '2000-01-01';
                    nextYearEndDay = currentDayWithNumberOfDaysToRetrieve - 31;
                    nextYearEndDate = '2000-01-' + nextYearEndDay;
                    filter = "BirthDate ge '" + today + "' or (BirthDate ge '" + nextYearStartDate + "' and BirthDate le '" + nextYearEndDate + "')";
                }
            }

            const usersSharePoint: any[] = await sp.web
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
                //.usingCaching()
                .get();

            if (usersSharePoint && usersSharePoint.length > 0) {

                allBirthDays = UserInformationMapper.mapToUserInformations(usersSharePoint);

                // First, select all bithdays of December to sort
                // Then, select all birthdays of the remaining months
                // Finally, contact both arrays and return
                if (currentMonth === '12') {
                    decemberBirthdays = allBirthDays.filter(b => moment(b.birthDate).format('MM') === '12');
                    decemberBirthdays = this.SortBirthdaysByBirthDate(decemberBirthdays);
                    monthsExceptoDecemberBirthdays = allBirthDays.filter(b =>moment(b.birthDate).format('MM') !== '12');
                    monthsExceptoDecemberBirthdays = this.SortBirthdaysByBirthDate(monthsExceptoDecemberBirthdays);
                    // Join the 2 arrays
                    allBirthDays = decemberBirthdays.concat(monthsExceptoDecemberBirthdays);
                }
                else {
                    allBirthDays = this.SortBirthdaysByBirthDate(allBirthDays);
                }
            }

            return allBirthDays;
        } catch (error) {
            Log.error(LOG_SOURCE, error, this._context.serviceScope);
            throw new Error(error.message);
        }
    }
}