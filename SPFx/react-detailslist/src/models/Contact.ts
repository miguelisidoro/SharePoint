import { IContact } from "./IContact"

export class Contact implements IContact
{
    Id: string;
    Name: string;
    Email: string;
    MobileNumber: string;

    constructor(obj: Partial<Contact> = {}) {
        this.Name = obj.Name;
        this.Email = obj.Email;
        this.MobileNumber = obj.MobileNumber;

        // Object.assign(this, obj);
    }
}