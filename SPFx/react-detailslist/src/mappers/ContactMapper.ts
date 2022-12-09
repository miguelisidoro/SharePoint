import { Contact, IContact, IContactSharePoint } from "../models";

export class ContactMapper {
    // Maps the list of contacts from SharePoint into an array of Contact model class
    static MapToContacts(contacts: IContactSharePoint[]): IContact[] {
        let mappedContacts: Contact[] = new Array<Contact>();

        contacts.forEach(contactSharePoint => {
            let mappedContact: Contact = new Contact({
                Name: contactSharePoint.Title,
                Email: contactSharePoint.Email,
                MobileNumber: contactSharePoint.Telemovel
            })

            mappedContact.Id = contactSharePoint.Id;

            mappedContacts.push(mappedContact);
        });

        // const mappedContacts = contacts.map(contactSharePoint => 
        //     new Contact({
        //         Id: contactSharePoint.Id,
        //         Name: contactSharePoint.Title,
        //         Email: contactSharePoint.Email,
        //         MobileNumber: contactSharePoint.Telemovel
        //     }));

        return mappedContacts;
    }

    // Maps the contact from SharePoint into the Contact model
    static MapToContact(contactSharePoint: IContactSharePoint): IContact {
        let mappedContact = new Contact({
            Name: contactSharePoint.Title,
            Email: contactSharePoint.Email,
            MobileNumber: contactSharePoint.Telemovel
        });

        mappedContact.Id = contactSharePoint.Id;

        return mappedContact;
    }
}