import { Microsoft365Group } from "../../../data";

export interface IMyTeamsGraphState
{
    microsoft365Groups: Microsoft365Group[];
    microsoftGroupOptions: any[];
    selectedGroupId: string;
}