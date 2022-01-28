import { Microsoft365Group, PersonInformation } from "../../../data";

export interface IMyTeamsGraphState
{
    microsoft365Groups: Microsoft365Group[];
    microsoftGroupOptions: any[];
    selectedGroupMembers: PersonInformation[];
}