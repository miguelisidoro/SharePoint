export class UserInformation {
    public Title: string;
    public Email: string;
    public JobTitle: string;
    public BirthDate: Date;
    public HireDate: Date;
    public WorkAnniversary: Date;

    constructor(obj: Partial<UserInformation> = {}) {
        Object.assign(this, obj);
    }
}