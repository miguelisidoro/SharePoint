import { Log } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/graph/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SharePointFieldNames } from "../constants";
import { UserInformationMapper } from "../mappers";
import { UserInformation } from "../models";
import * as moment from 'moment';
import { InformationDisplayType, InformationType } from "../enums";
import * as localforage from "localforage";

const LOG_SOURCE: string = "BirthdaysWorkAnniverariesNewHires";
export class SharePointServiceProvider {
    private _sharePointRelativeListUrl: string;
    private _numberOfItemsToShow: number;
    private _numberOfDaysToRetrieve: number;

    constructor(private context: WebPartContext,
        sharePointRelativeListUrl: string,
        numberOfItemsToShow: number,
        numberOfDaysToRetrieve: number) {

        // Setup Context to PnP JS
        sp.setup({
            spfxContext: this.context
        });

        localforage.config({
            driver: localforage.INDEXEDDB,
            name: 'BirthdaysWorkAnniverariesNewCollaborators',
            version: 1.0,
            storeName: 'BirthdaysWorkAnniverariesNewCollaborators_Store',
            description: 'BirthdaysWorkAnniverariesNewCollaborators storage'
        });

        this._sharePointRelativeListUrl = sharePointRelativeListUrl;
        this._numberOfItemsToShow = numberOfItemsToShow;
        this._numberOfDaysToRetrieve = numberOfDaysToRetrieve;
    }

    // Sort birthdays by birthdate
    private sortUsersByBirthDate(users: UserInformation[]) {
        return users.sort((a, b) => {
            if (a.BirthDate > b.BirthDate) {
                return 1;
            }
            if (a.BirthDate < b.BirthDate) {
                return -1;
            }
            return 0;
        });
    }

    private sortUsersByHireDateAscending(users: UserInformation[]) {
        return users.sort((a, b) => {
            if (a.HireDate > b.HireDate) {
                return 1;
            }
            if (a.HireDate < b.HireDate) {
                return -1;
            }
            return 0;
        });
    }

    private sortUsersByHireDateDescending(users: UserInformation[]) {
        return users.sort((a, b) => {
            if (a.HireDate < b.HireDate) {
                return 1;
            }
            if (a.HireDate > b.HireDate) {
                return -1;
            }
            return 0;
        });
    }

    private sortUsers(users: UserInformation[], informationType: InformationType): UserInformation[] {
        if (informationType === InformationType.Birthdays) {
            users = this.sortUsersByBirthDate(users);
        }
        else if (informationType === InformationType.WorkAnniversaries) {
            users = this.sortUsersByHireDateAscending(users);
        }
        else {
            users = this.sortUsersByHireDateDescending(users);
        }

        return users;
    }

    public getBirthdaysWorkAnniversariesViewXml(filterField: string, beginDate: string, endDate: string, rowLimit: number): string {
        const viewXml = `<View Scope='RecursiveAll'>
        <Query>
            <Where>
            <And>
                <Geq>
                    <FieldRef Name='${filterField}' />
                    <Value IncludeTimeValue='TRUE' Type='DateTime'>${beginDate}</Value>
                </Geq>
                <Lt>
                    <FieldRef Name='${filterField}' />
                    <Value IncludeTimeValue='TRUE' Type='DateTime'>${endDate}</Value>
                </Lt>
            </And>
            </Where>
            <OrderBy>
            <FieldRef Name='${filterField}' Ascending='True' />
            </OrderBy>
        </Query>
        <ViewFields><FieldRef Name='${SharePointFieldNames.Id}'/><FieldRef Name='${SharePointFieldNames.Title}'/><FieldRef Name='${SharePointFieldNames.Email}'/><FieldRef Name='${SharePointFieldNames.JobTitle}'/><FieldRef Name='${SharePointFieldNames.BirthDate}'/><FieldRef Name='${SharePointFieldNames.HireDate}'/></ViewFields>
        <RowLimit Paged='TRUE'>${rowLimit}</RowLimit></View>`;

        return viewXml;
    }

    // Get users anniversaries or new collaborators
    // Important NOTE: All dates are stored with year 2000
    public async getAnniversariesOrNewCollaborators(
        informationType: InformationType,
        informationDisplayType: InformationDisplayType): Promise<UserInformation[]> {
        let currentDate: Date, today: string, currentMonth: string, currentDay: number;
        let filter: string;
        let otherYearUsers: UserInformation[], currentYearUsers: UserInformation[], allUsers: UserInformation[];

        try {
            const filterField: string = informationType === InformationType.Birthdays ?
                SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;

            // Build filter to use in query to SharePoint
            today = '2000-' + moment().format('MM-DD');
            //today = '2000-04-12';
            currentMonth = moment().format('MM');
            //currentMonth = '04';
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

            // Retrieve users from SharePoint
            const usersSharePoint = await sp.web
                .getList(this._sharePointRelativeListUrl)
                .items.select(
                    SharePointFieldNames.Id,
                    SharePointFieldNames.Title,
                    SharePointFieldNames.Email,
                    SharePointFieldNames.JobTitle,
                    SharePointFieldNames.BirthDate,
                    SharePointFieldNames.HireDate)
                .filter(filter)
                .top(5000) //avoid SharePoint List View Threshold, we will return the desired number of users in the end
                .usingCaching()
                .get();

            // Sort users
            if (usersSharePoint && usersSharePoint.length > 0) {

                allUsers = UserInformationMapper.mapToUserInformations(usersSharePoint);

                if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                    // If end date is from next year, first, select all anniversaries from current year to sort
                    // Then, select all anniversaries of the remaining months from next year
                    // Finally, contact both arrays and return
                    if (currentDatewithDaysToRetrieveYear === '2001') {
                        // get the anniversaries from the current year (months are >= currentMonth)
                        if (informationType === InformationType.Birthdays) {
                            currentYearUsers = allUsers.filter(b => moment(b.BirthDate).month() + 1 >= parseInt(currentMonth));
                        }
                        else {
                            currentYearUsers = allUsers.filter(b => moment(b.HireDate).month() + 1 >= parseInt(currentMonth));
                        }
                        currentYearUsers = this.sortUsers(currentYearUsers, informationType);
                        if (informationType === InformationType.Birthdays) {
                            otherYearUsers = allUsers.filter(b => moment(b.BirthDate).month() + 1 < parseInt(currentMonth));
                        }
                        else {
                            otherYearUsers = allUsers.filter(b => moment(b.HireDate).month() + 1 < parseInt(currentMonth));
                        }
                        otherYearUsers = this.sortUsers(otherYearUsers, informationType);
                        // Join the 2 arrays
                        allUsers = currentYearUsers.concat(otherYearUsers);
                    }
                    else {
                        allUsers = this.sortUsers(allUsers, informationType);
                    }
                }
                else {
                    // New Hires
                    // If end date is from previous year, first, select all new hires from current year to sort
                    // Then, select all of the remaining months from previous year
                    // Finally, contact both arrays, starting by previous year users and return
                    if (currentDatewithDaysToRetrieveYear === '1999') {
                        // get the new hires from the previous year (months are >= currentMonth)
                        currentYearUsers = allUsers.filter(b => moment(b.HireDate).month() + 1 <= parseInt(currentMonth));
                        currentYearUsers = this.sortUsers(currentYearUsers, informationType);
                        otherYearUsers = allUsers.filter(b => moment(b.HireDate).month() + 1 > parseInt(currentMonth));
                        otherYearUsers = this.sortUsers(otherYearUsers, informationType);
                        // Join the 2 arrays
                        allUsers = currentYearUsers.concat(otherYearUsers);
                    }
                    else {
                        allUsers = this.sortUsers(allUsers, informationType);
                    }
                }
            }

            let users: UserInformation[];

            if (informationDisplayType === InformationDisplayType.TopResults) {
                // If we are in Top Results mode, we want to get only the top results from the cached data
                // Filter is done in the end so that we can get all the users for the number of days to retrieve and then from those users, return the number of items to show
                users = allUsers.slice(0, this._numberOfItemsToShow - 1);
            }
            // else //More results
            // {
            //     //if we are in more results, get the current page
            //     users = allUsers.slice(skip, take);
            // }

            return users;
        } catch (error) {
            Log.error(LOG_SOURCE, error, this.context.serviceScope);
            throw new Error(error.message);
        }
    }

    public async getAnniversariesOrNewCollaboratorsRenderListDataAsStream(
        informationType: InformationType,
        informationDisplayType: InformationDisplayType): Promise<UserInformation[]> {
        {
            const cacheKey = 'AnniversariesOrNewCollaborators';

            let cachedAnniversariesOrNewCollaborators: UserInformation[] = await localforage.getItem(cacheKey);

            if (cachedAnniversariesOrNewCollaborators !== null && cachedAnniversariesOrNewCollaborators !== undefined && cachedAnniversariesOrNewCollaborators.length > 0) {
                return cachedAnniversariesOrNewCollaborators;
            }
            try {
                const today = '2000-' + moment().format('MM-DD');
                const currentDate = moment(today).toDate();
                const currentDatewithDaysToRetrieve = currentDate;
                currentDatewithDaysToRetrieve.setDate(currentDate.getDate() + this._numberOfDaysToRetrieve);

                const currentDateMidNight = '2000-' + moment().format('MM-DD') + 'T00:00:00Z';
                const currentDatewithDaysToRetrieveMidNight = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD') + 'T00:00:00Z';

                const filterField = informationType === InformationType.Birthdays ? SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;

                let viewXml = this.getBirthdaysWorkAnniversariesViewXml(
                    filterField,
                    currentDateMidNight,
                    currentDatewithDaysToRetrieveMidNight,
                    this._numberOfItemsToShow);

                const usersSharePointCurrentYear = await sp.web.getList(this._sharePointRelativeListUrl).renderListDataAsStream({
                    ViewXml: viewXml
                });

                // check if we have enough data to display with dates from current year
                if (usersSharePointCurrentYear !== null && usersSharePointCurrentYear !== undefined && usersSharePointCurrentYear.Row !== null && usersSharePointCurrentYear.Row !== undefined && usersSharePointCurrentYear.Row.length === this._numberOfItemsToShow) {
                    const mappedUsersSharePoint = UserInformationMapper.mapToUserInformations(usersSharePointCurrentYear.Row);

                    localforage.setItem(cacheKey, mappedUsersSharePoint);

                    return mappedUsersSharePoint;
                }
                else {
                    // if not, get dates from the next year
                    viewXml = this.getBirthdaysWorkAnniversariesViewXml(
                        filterField,
                        '2000-01-01T00:00:00Z',
                        currentDateMidNight,
                        this._numberOfItemsToShow - usersSharePointCurrentYear.Row.length);

                    const usersSharePointNextYear = await sp.web.getList(this._sharePointRelativeListUrl).renderListDataAsStream({
                        ViewXml: viewXml
                    });

                    const mappedUsersSharePointCurrentYear = UserInformationMapper.mapToUserInformations(usersSharePointCurrentYear.Row);

                    const mappedUsersSharePointNextYear = UserInformationMapper.mapToUserInformations(usersSharePointNextYear.Row);

                    const mappedUsersSharePoint = mappedUsersSharePointCurrentYear.concat(mappedUsersSharePointNextYear);

                    localforage.setItem(cacheKey, mappedUsersSharePoint);

                    return mappedUsersSharePoint;
                }
            }
            catch (error) {
                Log.error(LOG_SOURCE, error, this.context.serviceScope);
                throw new Error(error.message);
            }
        }
    }
}