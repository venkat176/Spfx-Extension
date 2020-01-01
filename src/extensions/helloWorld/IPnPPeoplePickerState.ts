import { DayOfWeek } from 'office-ui-fabric-react/lib/DatePicker';
import { IDropdownOption } from 'office-ui-fabric-react';

export interface IPnPPeoplePickerState {  
    addUsers: string[];  
    selectedItem?: IDropdownOption;
    selectedITem?: IDropdownOption;
    firstDayOfWeek?: DayOfWeek;
    value?: Date | null;
    saving: boolean;
    date:Date;
    title:string;
    dpvalue:string;
    dropvalue:string;
    dudate:Date,
}  