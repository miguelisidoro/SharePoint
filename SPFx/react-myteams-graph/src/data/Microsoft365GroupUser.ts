export class Microsoft365GroupUser
{
    name: string;
    email: string;
    userPrincipalName: string;
    jobTitle: string;
    
    constructor(obj: Partial<Microsoft365GroupUser> = {}) {
        this.name = obj.name;
        this.email = obj.email;
        this.userPrincipalName = obj.userPrincipalName;
        this.jobTitle = obj.jobTitle;
    }
}