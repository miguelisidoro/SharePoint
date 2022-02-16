export class UserInformation
{
    Title: string;
    Email: string;
    JobTitle: string;
    BirthDate: Date;
    HireDate: Date;

    constructor(obj: Partial<UserInformation> = {}) {
        Object.assign(this, obj)
    }
}