export class Collection
{
    name?: string;
    description?: string;
    imageUrl?: string;
    slug?: string;

    constructor(obj: Partial<Collection> = {}) {
        Object.assign(this, obj);
    }
}