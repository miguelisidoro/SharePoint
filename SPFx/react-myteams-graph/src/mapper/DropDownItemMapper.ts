import { IDropdownOption } from '@fluentui/react';
import { Microsoft365Group } from '../data';

export class DropDownItemMapper {
    public static MapToDropDownItems(microsoft365Groups: Microsoft365Group[]): any[] {
        let mappedDropDownItems: any[] = [];

        microsoft365Groups.forEach(microsoft365Group =>
            mappedDropDownItems.push({
                key: microsoft365Group.GroupId,
                text: microsoft365Group.GroupName
            }));

        return mappedDropDownItems;
    }
}