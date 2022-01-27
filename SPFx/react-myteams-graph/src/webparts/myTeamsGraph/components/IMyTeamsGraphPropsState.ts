import { Microsoft365Group } from "../../../data";

export interface IMyTeamsGraphState
{
    currentUserGroups: Microsoft365Group[];
}