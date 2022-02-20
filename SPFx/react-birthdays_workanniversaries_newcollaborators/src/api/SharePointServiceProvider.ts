import { Log } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/graph/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { CacheExpiration, SharePointFieldNames } from "../constants";
import { UserInformationMapper } from "../mappers";
import { UserInformation } from "../models";
import * as moment from 'moment';
import { InformationDisplayType, InformationType } from "../enums";
import * as localforage from "localforage";
import cache from "@luudjanssen/localforage-cache";

const LOG_SOURCE: string = "BirthdaysWorkAnniverariesNewHires";

const birthdaysWorkAnniversariesNewCollaboratorsCache = cache.createInstance({
    name: "BirthdaysWorkAnniversariesNewCollaboratorsCache",
    defaultExpiration: CacheExpiration.BirthdaysWorkAnniversariesNewCollaboratorsCacheExpiration
});

export class SharePointServiceProvider {
    private _sharePointRelativeListUrl: string;
    private _numberOfDaysToRetrieve: number;

    constructor(private context: WebPartContext,
        sharePointRelativeListUrl: string,
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
        this._numberOfDaysToRetrieve = numberOfDaysToRetrieve;
    }

    public getBirthdaysWorkAnniversariesNewCollaboratorsViewXml(informationType: InformationType, beginDate: string, endDate: string, rowLimit: number): string {
        let filterField;
        if (informationType === InformationType.Birthdays) {
            filterField = SharePointFieldNames.BirthDate;
        }
        else if (informationType === InformationType.WorkAnniversaries) {
            filterField = SharePointFieldNames.WorkAnniversary;
        }
        else { // New Collaborators
            filterField = SharePointFieldNames.HireDate;
        }
        const sortAscending = (informationType === InformationType.Birthdays || informationType === InformationType.WorkAnniversaries) ?
            true : false;

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
    public async getAnniversariesOrNewCollaborators(
        informationType: InformationType,
        informationDisplayType: InformationDisplayType,
        rowLimit: number): Promise<UserInformation[]> {
        {
            try {
                let cacheKey = InformationType[informationType] + "Cache";

                //check if users are in cache and return from cache if they are

                let cachedAnniversariesOrNewCollaborators: UserInformation[] = await birthdaysWorkAnniversariesNewCollaboratorsCache.getItem(cacheKey);

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
                else {
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

                    if (numberOfDaysBetweenCurrentDateAndBeginningOfYear >= 0 && numberOfDaysBetweenCurrentDateAndBeginningOfYear < this._numberOfDaysToRetrieve) {
                        beginDate = currentDatewithDaysToRetrieveMidNight;
                    }
                    else {
                        beginDate = '2000-01-01T00:00:00Z';
                    }
                    endDate = currentDateMidNight;
                }

                // get CAML Query to call SharePoint
                let viewXml = this.getBirthdaysWorkAnniversariesNewCollaboratorsViewXml(
                    informationType,
                    beginDate,
                    endDate,
                    rowLimit);

                //get data from SharePoint
                const usersSharePointCurrentYear = await sp.web.getList(this._sharePointRelativeListUrl).renderListDataAsStream({
                    ViewXml: viewXml
                });

                // check if we have enough data to display with dates from current year
                if (usersSharePointCurrentYear !== null && usersSharePointCurrentYear !== undefined && usersSharePointCurrentYear.Row !== null && usersSharePointCurrentYear.Row !== undefined && usersSharePointCurrentYear.Row.length === rowLimit) {
                    //if there is enough data, map into the object we want to return
                    const mappedUsersSharePoint = UserInformationMapper.mapToUserInformations(usersSharePointCurrentYear.Row);

                    //store data in cache
                    birthdaysWorkAnniversariesNewCollaboratorsCache.setItem(cacheKey, mappedUsersSharePoint);

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
                        rowLimit - usersSharePointCurrentYear.Row.length);

                    const usersSharePointNextYear = await sp.web.getList(this._sharePointRelativeListUrl).renderListDataAsStream({
                        ViewXml: viewXml
                    });

                    const mappedUsersSharePointCurrentYear = UserInformationMapper.mapToUserInformations(usersSharePointCurrentYear.Row);

                    const mappedUsersSharePointOtherYear = UserInformationMapper.mapToUserInformations(usersSharePointNextYear.Row);

                    // concat the data from current year and the other year
                    const mappedUsersSharePoint = mappedUsersSharePointCurrentYear.concat(mappedUsersSharePointOtherYear);

                    //store data in cache
                    birthdaysWorkAnniversariesNewCollaboratorsCache.setItem(cacheKey, mappedUsersSharePoint);

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