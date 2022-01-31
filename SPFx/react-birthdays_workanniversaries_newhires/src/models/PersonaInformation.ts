export class PersonaInformation
{
    imageUrl: string;
    text: string;
    secondaryText: string;
    userPrincipalName: string;

    constructor(obj: Partial<PersonaInformation> = {}) {
        this.imageUrl = obj.imageUrl;
        this.text = obj.text;
        this.secondaryText = obj.secondaryText;
        this.userPrincipalName = obj.userPrincipalName;
    }
}