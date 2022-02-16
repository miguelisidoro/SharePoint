export class Collection
{
    name?: string;
    description?: string;
    imageUrl?: string;

    constructor(obj: Partial<Collection> = {}) {
        Object.assign(this, obj);
    }
}