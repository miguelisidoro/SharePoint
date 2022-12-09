export class PersonaInformation {
    public imageUrl: string;
    public text: string;
    public secondaryText: string;
    public userPrincipalName: string;
    public isAnniversaryToday: boolean;

    constructor(obj: Partial<PersonaInformation> = {}) {
        Object.assign(this, obj);
    }
}