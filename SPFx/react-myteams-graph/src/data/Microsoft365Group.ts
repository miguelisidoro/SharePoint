export class Microsoft365Group
{
    GroupId: string;
    GroupName: string;

    constructor(obj: Partial<Microsoft365Group> = {}) {
        this.GroupId = obj.GroupId;
        this.GroupName = obj.GroupName;
    }
}