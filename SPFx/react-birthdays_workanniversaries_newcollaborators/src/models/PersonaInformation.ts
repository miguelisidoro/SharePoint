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
}