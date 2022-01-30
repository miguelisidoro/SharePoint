export class UserInformation
{
    title: string;
    userTitle: string;
    userEmail: string;
    jobTitle: string;
    birthDate: Date;
    hireDate: Date;

    constructor(obj: Partial<UserInformation> = {}) {
        this.title = obj.title;
        this.userTitle = obj.userTitle;
        this.userEmail = obj.userEmail;
        this.jobTitle = obj.jobTitle;
        this.birthDate = obj.birthDate;
        this.hireDate = obj.hireDate;
    }
}