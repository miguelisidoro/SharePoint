import { Contact, IContact, IContactSharePoint } from "../models";

export class ContactMapper
{
    static MapToContact(contacts : IContactSharePoint[]) : IContact[]
    {
        debugger;
        const mappedContacts = contacts.map(contactSharePoint => 
            new Contact({
                Name: contactSharePoint.Title,
                Email: contactSharePoint.Email,
                MobileNumber: contactSharePoint.Telemovel
            }));

        return mappedContacts;
    }
}