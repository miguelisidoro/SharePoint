export class PersonInformation
{
    imageUrl: string;
    text: string;

    constructor(obj: Partial<PersonInformation> = {}) {
        this.imageUrl = obj.imageUrl;
        this.text = obj.text;
    }
}

