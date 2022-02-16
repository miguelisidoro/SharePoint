export class PersonaInformation
{
    imageUrl: string;
    text: string;
    secondaryText: string;
    userPrincipalName: string;
    isAnniversaryToday: boolean;


    constructor(obj: Partial<PersonaInformation> = {}) {
        Object.assign(this, obj)
    }

    // constructor(obj: Partial<PersonaInformation> = {}) {
    //     Object.assign(this, obj);
    //     this.imageUrl = obj.imageUrl;
    //     this.text = obj.text;
    //     this.secondaryText = obj.secondaryText;
    //     this.userPrincipalName = obj.userPrincipalName;
    //     this.isAnniversaryToday = obj.isAnniversaryToday;
    // }
}