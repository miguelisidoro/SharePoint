import { Contact } from "../models/Contact";
import { IContact } from "../models/IContact";
import { IContactSharePoint } from "../models/IContactSharePoint";

export class ContactMapper
{
    static MapToContact(contacts : IContactSharePoint[]) : IContact[]
    {
        const mappedContacts = contacts.map(contactSharePoint => 
            new Contact({
                Name = contactSharePoint.Title,
                Email = contactSharePoint.Email,
                MobileNumber = contactSharePoint.Telemovel
            }));

        return mappedContacts;
    }
}