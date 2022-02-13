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
            storeName: 'BirthdaysWorkAnniverariesNewCollaborators',
            description: 'Birthdays, Work Anniveraries, New Collaborators Indexed DB Storage'
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

    public getBirthdaysWorkAnniversariesNewCollaboratorsViewXml(informationType: InformationType, beginDate: string, endDate: string, rowLimit: number): string {
        const filterField: string = informationType === InformationType.Birthdays ?
        SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;

        const sortAscending = (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) ?
        true : false;

        // const queryCamlFirstOperator = (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) ?
        // 'Geq' : 'Gt';

        const queryCamlSecondOperator = (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) ?
        'Lt' : 'Leq';

        const viewXml = `<View Scope='RecursiveAll'>
        <Query>
            <Where>
            <And>
                <Geq>
                    <FieldRef Name='${filterField}' />
                    <Value IncludeTimeValue='TRUE' Type='DateTime'>${beginDate}</Value>
                </Geq>
                <${queryCamlSecondOperator}>
                    <FieldRef Name='${filterField}' />
                    <Value IncludeTimeValue='TRUE' Type='DateTime'>${endDate}</Value>
                </${queryCamlSecondOperator}>
            </And>
            </Where>
            <OrderBy>
            <FieldRef Name='${filterField}' Ascending='${sortAscending.toString().toUpperCase()}' />
            </OrderBy>
        </Query>
        <ViewFields><FieldRef Name='${SharePointFieldNames.Id}'/><FieldRef Name='${SharePointFieldNames.Title}'/><FieldRef Name='${SharePointFieldNames.Email}'/><FieldRef Name='${SharePointFieldNames.JobTitle}'/><FieldRef Name='${SharePointFieldNames.BirthDate}'/><FieldRef Name='${SharePointFieldNames.HireDate}'/></ViewFields>
        <RowLimit Paged='TRUE'>${rowLimit}</RowLimit></View>`;

        return viewXml;
    }

    // Get users anniversaries or new collaborators
    // Important NOTE: All dates are stored with year 2000
    //TODO: remove
    public async getAnniversariesOrNewCollaboratorsOld(
        informationType: InformationType,
        informationDisplayType: InformationDisplayType): Promise<UserInformation[]> {
        let currentDate: Date, today: string, currentMonth: string, currentDay: number;
        let filter: string;
        let otherYearUsers: UserInformation[], currentYearUsers: UserInformation[], allUsers: UserInformation[];

        const filterField: string = informationType === InformationType.Birthdays ?
        SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;

        try {
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
                // If end date is from previous year, get new hires from both years
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
                    // New Collaborators
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

    // Get users anniversaries or new collaborators
    // Important NOTE: All dates are stored with year 2000
    public async getAnniversariesOrNewCollaborators(
        informationType: InformationType,
        informationDisplayType: InformationDisplayType): Promise<UserInformation[]> {
        {
            try {
                let cacheKey = InformationType[informationType];

                //check if users are in cache and return from cache if they are
                let cachedAnniversariesOrNewCollaborators: UserInformation[] = await localforage.getItem(cacheKey);

                if (cachedAnniversariesOrNewCollaborators !== null && cachedAnniversariesOrNewCollaborators !== undefined && cachedAnniversariesOrNewCollaborators.length > 0) {
                    return cachedAnniversariesOrNewCollaborators;
                }

                let beginDate, endDate;
                const today = '2000-' + moment().format('MM-DD');
                //const today = '2000-01-07';
                const currentDate = moment(today).toDate();
                const currentDatewithDaysToRetrieve = currentDate;

                //get date considering number of days to retrieve
                //we cannot retrieve the whole year due to performance if we have a lot of users
                if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                    currentDatewithDaysToRetrieve.setDate(currentDate.getDate() + this._numberOfDaysToRetrieve);
                }
                else
                {
                    currentDatewithDaysToRetrieve.setDate(currentDate.getDate() - this._numberOfDaysToRetrieve);
                }

                const currentDateMidNight = '2000-' + moment(today).format('MM-DD') + 'T00:00:00Z';
                const currentDatewithDaysToRetrieveMidNight = '2000-' + moment(currentDatewithDaysToRetrieve).format('MM-DD') + 'T00:00:00Z';

                //get begin and end dates to filter data
                if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                    beginDate = currentDateMidNight;
                
                    // check if end date should be the end of year or the current date plus days to retrieve depending on the difference between current date and end of year
                    const endOfYearDate = moment('2000-12-31').toDate();

                    let numberOfMiliSecondsBetweenCurrentDateAndEndOfYear: number = endOfYearDate.getTime() - currentDate.getTime();
                    let numberOfDaysBetweenCurrentDateAndEndOfYear: number = Math.ceil(numberOfMiliSecondsBetweenCurrentDateAndEndOfYear / (1000 * 60 * 60 * 24));

                    if (numberOfDaysBetweenCurrentDateAndEndOfYear < this._numberOfDaysToRetrieve) {
                        endDate = '2000-12-31T00:00:00Z';
                    }
                    else {
                        endDate = currentDatewithDaysToRetrieveMidNight;
                    }
                }
                else //New Collaborators
                {
                    // check if begin date should be the beginning of year or the current date minus days to retrieve depending on the difference between current date and beginning of year
                    const begginingOfYearDate = moment('2000-01-01').toDate();

                    let numberOfMiliSecondsBetweenCurrentDateAndBeginningOfYear: number = currentDate.getTime() - begginingOfYearDate.getTime();
                    let numberOfDaysBetweenCurrentDateAndBeginningOfYear: number = Math.ceil(numberOfMiliSecondsBetweenCurrentDateAndBeginningOfYear / (1000 * 60 * 60 * 24));

                    if (numberOfDaysBetweenCurrentDateAndBeginningOfYear >= 0 &&numberOfDaysBetweenCurrentDateAndBeginningOfYear < this._numberOfDaysToRetrieve) {
                        beginDate = currentDatewithDaysToRetrieveMidNight;
                    }
                    else {
                        beginDate = '2000-01-01T00:00:00Z';
                    }
                    endDate = currentDateMidNight;
                }

                // get CAML Query to call SharePoint
                const filterField = informationType === InformationType.Birthdays ? SharePointFieldNames.BirthDate : SharePointFieldNames.HireDate;

                let viewXml = this.getBirthdaysWorkAnniversariesNewCollaboratorsViewXml(
                    informationType,
                    beginDate,
                    endDate,
                    this._numberOfItemsToShow);

                //get data from SharePoint
                const usersSharePointCurrentYear = await sp.web.getList(this._sharePointRelativeListUrl).renderListDataAsStream({
                    ViewXml: viewXml
                });

                // check if we have enough data to display with dates from current year
                if (usersSharePointCurrentYear !== null && usersSharePointCurrentYear !== undefined && usersSharePointCurrentYear.Row !== null && usersSharePointCurrentYear.Row !== undefined && usersSharePointCurrentYear.Row.length === this._numberOfItemsToShow) {
                    //if there is enough data, map into the object we want to return
                    const mappedUsersSharePoint = UserInformationMapper.mapToUserInformations(usersSharePointCurrentYear.Row);

                    //store data in cache
                    localforage.setItem(cacheKey, mappedUsersSharePoint);

                    return mappedUsersSharePoint;
                }
                else {
                    //if we don't have enough data, get data from othe year (next year if birthdays or work anniversaries or previous year if new collaborators)
                    if (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) {
                        beginDate = '2000-01-01T00:00:00Z';
                        endDate = currentDateMidNight;
                    }
                    else //New Collaborators
                    {
                        beginDate = '2000-' + moment(currentDatewithDaysToRetrieveMidNight).format('MM-DD') + 'T00:00:00Z';
                        endDate = '2000-12-31T00:00:00Z';
                    }
                    viewXml = this.getBirthdaysWorkAnniversariesNewCollaboratorsViewXml(
                        informationType,
                        beginDate,
                        endDate,
                        this._numberOfItemsToShow - usersSharePointCurrentYear.Row.length);

                    const usersSharePointNextYear = await sp.web.getList(this._sharePointRelativeListUrl).renderListDataAsStream({
                        ViewXml: viewXml
                    });

                    const mappedUsersSharePointCurrentYear = UserInformationMapper.mapToUserInformations(usersSharePointCurrentYear.Row);

                    const mappedUsersSharePointOtherYear = UserInformationMapper.mapToUserInformations(usersSharePointNextYear.Row);

                    // concat the data from current year and the other year
                    const mappedUsersSharePoint = mappedUsersSharePointCurrentYear.concat(mappedUsersSharePointOtherYear);

                    //store data in cache
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