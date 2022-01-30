export class UserInformation
{
    imageUrl: string;
    text: string;
    secondaryText: string;
    userPrincipalName: string;

    constructor(obj: Partial<UserInformation> = {}) {
        this.imageUrl = obj.imageUrl;
        this.text = obj.text;
        this.secondaryText = obj.secondaryText;
        this.userPrincipalName = obj.userPrincipalName;
    }
}