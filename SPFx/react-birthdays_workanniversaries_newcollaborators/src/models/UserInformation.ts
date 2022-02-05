export class UserInformation
{
    Title: string;
    Email: string;
    JobTitle: string;
    BirthDate: Date;
    HireDate: Date;

    constructor(obj: Partial<UserInformation> = {}) {
        this.Title = obj.Title;
        this.Email = obj.Email;
        this.JobTitle = obj.JobTitle;
        this.BirthDate = obj.BirthDate;
        this.HireDate = obj.HireDate;
    }
}