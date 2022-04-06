import { UserInformation } from ".";

export class PagedUserInformation {
    public users: UserInformation[];
    public nextPageUrl: string;

    constructor(obj: Partial<PagedUserInformation> = {}) {
        Object.assign(this, obj);
    }
}