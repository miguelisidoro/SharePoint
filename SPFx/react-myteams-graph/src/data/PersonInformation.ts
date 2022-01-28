export class PersonInformation
{
    imageUrl: string;
    text: string;
    secondaryText: string;

    constructor(obj: Partial<PersonInformation> = {}) {
        this.imageUrl = obj.imageUrl;
        this.text = obj.text;
        this.secondaryText = obj.secondaryText;
    }
}

