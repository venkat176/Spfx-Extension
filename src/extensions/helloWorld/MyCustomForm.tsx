import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp,Web } from 'sp-pnp-js';
// import { Web } from '@pnp/sp';
// import pnp, { sp, Item, ItemAddResult, ItemUpdateResult,Web } from "sp-pnp-js";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { autobind } from 'office-ui-fabric-react';
import { getGUID } from "@pnp/common";
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { PrimaryButton } from 'office-ui-fabric-react';
import * as pnp from 'sp-pnp-js';
import $ from "jquery";
import { IPnPPeoplePickerProps } from './IPnPPeoplePickerProps';
import { IPnPPeoplePickerState } from './IPnPPeoplePickerState';
import { SPHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base'; 

// import { useConstCallback } from '@uifabric/react-hooks';

const DayPickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker'
};

//   export interface IDatePickerBasicExampleState {
//     firstDayOfWeek?: DayOfWeek;
//   }

const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px'
    }
});

// export interface ISPList {
//     ID: string;
//     Title: string;
//     Status: string;
//     Priority: string;
//     DueDate: Date;
// }

 class MycustomForm extends React.Component<IPnPPeoplePickerProps, IPnPPeoplePickerState> {
private _context =WebPartContext; 

    constructor(props: IPnPPeoplePickerProps) {
        super(props);
        this.state = {
            value: null,
            firstDayOfWeek: DayOfWeek.Monday,
            addUsers: [],
            saving: false,
            date:null,
            issuestatus:[],
            title:""
        };
    }

    public render(): React.ReactElement<IPnPPeoplePickerProps> {
        const { selectedItem } = this.state;
        const { selectedITem } = this.state;
        const { firstDayOfWeek, value } = this.state;

        return (
            <div>
                <h2>New Item</h2>
                
                <Label htmlFor="title" required>Title</Label>
                <TextField id="title" ariaLabel="text field" onChange={this.issueName} />
                <Label htmlFor="assignedto">Assigned To</Label>
                <PeoplePicker
                    context={this.context}
                    titleText="People Picker"
                    personSelectionLimit={3}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this._getPeoplePickerItems}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                <Label htmlFor="issuestatus">Issue Status</Label>
                <Dropdown       
                    id="issuestatus"
                    selectedKey={selectedItem ? selectedItem.key : undefined}
                    onChange={this._onChange}
                    placeholder="Select an option"
                    options={[
                        { key: 'Active', text: 'Active' },
                        { key: 'Resolved', text: 'Resolved' },
                        { key: 'Closed', text: 'Closed' },
                    ]}
                // defaultValue  ={this.state.selectedItem}
                // styles={{ dropdown: { width: 300 } }}
                />
                <Label htmlFor="priority">Priority</Label>
                <Dropdown
                    // label="Controlled example"
                    selectedKey={selectedITem ? selectedITem.key : undefined}
                    id="priority"
                    onChange={this.onchange}
                    placeholder="Select an option"
                    options={[
                        { key: '(1)High', text: '(1)High' },
                        { key: '(2)Normal', text: '(2)Normal' },
                        { key: '(3)Low', text: '(3)Low' },
                    ]} />
                <Label htmlFor="duedate">Due Date</Label>
                <DatePicker
                    className={controlClass.control}
                    firstDayOfWeek={firstDayOfWeek}
                    strings={DayPickerStrings}
                    showWeekNumbers={true}
                    firstWeekOfYear={1}
                    showMonthPickerAsOverlay={true}
                    formatDate={this._onFormatDate}
                    id="duedate"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    value={value!}
                    parseDateFromString={this._onParseDateFromString}
                    onSelectDate={this._onSelectDate}
                />
                {/* <DateTimePicker label="DateTime Picker - 24h"
                dateConvention={DateConvention.DateTime}
                timeConvention={TimeConvention.Hours24}
                value={this.state.dudate}
                onChange={this.handleChange} /> */}
                <PrimaryButton text="save" onClick={_alertClicked} />
                {/* <DefaultButton text="cancel" /> */}
            </div >

        );
        // this.AddEventListeners();
    }

    // @autobind
    // private addSelectedUsers(): void {
    //     sp.web.lists.getByTitle("SPFx Users").items.add({
    //         // Title: ,
    //         Users: {
    //             results: this.state.addUsers
    //         }
    //     }).then(i => {
    //         console.log(i);
    //     });
    // }
    // private _onSelectDate = (date: Date | null | undefined): void => {
    //     this.setState({ dudate: date });
    // };

    private _onSelectDate = (date: Date | null | undefined): void => {
        this.setState({ value: date });
    };
    private isssueName(issueTitle:string) {
        this.setState({ title: issueTitle });
    };
    private _getPeoplePickerItems(items: any[]) {
        console.log('Items:', items);
        this.setState({ addUsers: items });
    }
    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    };
    private _onParseDateFromString = (value: string): Date => {
        const date = this.state.value || new Date();
        const values = (value || '').trim().split('/');
        const day = values.length > 0 ? Math.max(1, Math.min(31, parseInt(values[0], 10))) : date.getDate();
        const month = values.length > 1 ? Math.max(1, Math.min(12, parseInt(values[1], 10))) - 1 : date.getMonth();
        let year = values.length > 2 ? parseInt(values[2], 10) : date.getFullYear();
        if (year < 100) {
            year += date.getFullYear() - (date.getFullYear() % 100);
        }
        return new Date(year, month, day);
    };
    private _onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        // console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
        this.setState({ selectedItem: item });
    };
    private onchange = (event: React.FormEvent<HTMLDivElement>, Item: IDropdownOption): void => {
        // console.log(`Selection change: ${Item.text} ${Item.selected ? 'selected' : 'unselected'}`);
        this.setState({ selectedITem: Item });
    };
}

function _alertClicked(): void {
    //alert('Clicked');
    pnp.sp.web.lists.getByTitle('issueTest').items.add({
        Title: document.getElementById('title')["value"],
        // AssignedTo : document.getElementById('assignedto')["value"],   
        Status: document.getElementById('issuestatus')["value"],
        Priority: document.getElementById('priority')["value"],
        DueDate: document.getElementById('duedate')["value"]
    });
    alert("Record with Profile Name : " + document.getElementById('title')["value"] + " Added !");
}