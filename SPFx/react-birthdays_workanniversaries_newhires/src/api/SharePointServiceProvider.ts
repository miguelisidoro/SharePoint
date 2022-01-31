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

    constructor(private _context: WebPartContext, sharePointRelativeListUrl: string) {
        // Setup Context to PnPjs and MSGraph
        sp.setup({
            spfxContext: this._context
        });

        this._sharePointRelativeListUrl = sharePointRelativeListUrl;

        this.onInit();
    }

    private async onInit() { }

    // Get users from User Information SharePoint List
    public async getUsers(): Promise<UserInformation[]> {
        try {
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
                .usingCaching()
                .orderBy(SharePointFieldNames.Title)
                .get();

            let mappedUsers = UserInformationMapper.mapToUserInformations(usersSharePoint);

            return mappedUsers;
        } catch (error) {
            Log.error(LOG_SOURCE, error, this._context.serviceScope);
            throw new Error(error.message);
        }

        
        // let _results, _today: string, _month: string, _day: number;
        // let _filter: string, _countdays: number, _f:number, _nextYearStart: string;
        // let  _FinalDate: string;
        // try {
        // _results = null;
        // _today = '2000-' + moment().format('MM-DD');
        // _month = moment().format('MM');
        // _day = parseInt(moment().format('DD'));
        // _filter = "fields/Birthday ge '" + _today + "'";
        // // If we are in Dezember we have to look if there are birthday in January
        // // we have to build a condition to select birthday in January based on number of upcommingDays
        // // we can not use the year for teste , the year is always 2000.
        // console.log(_month);
        // if (_month === '12') {
        //     _countdays = _day + upcommingDays;
        //     _f = 0;
        //     _nextYearStart = '2000-01-01';
        //     _FinalDate = '2000-01-';
        //     if ((_countdays) > 31) {
        //     _f = _countdays - 31;
        //     _FinalDate = _FinalDate + _f;
        //     _filter = "fields/Birthday ge '" + _today + "' or (fields/Birthday ge '" + _nextYearStart + "' and fields/Birthday le '" + _FinalDate + "')";
        //     }
        // }
        // this.graphClient = await this._context.msGraphClientFactory.getClient();
        // _results = await this.graphClient.api(`sites/root/lists('${this.birthdayListTitle}')/items?orderby=Fields/Birthday`)
        //     .version('v1.0')
        //     .expand('fields')
        //     .top(upcommingDays)
        //     .filter(_filter)
        //     .get();

        //     return _results.value;

        // } catch (error) {
        // console.dir(error);
        // return Promise.reject(error);
        // }
    }
}