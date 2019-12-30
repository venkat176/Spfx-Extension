import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  Panel,
  PanelType
} from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
//import MycustomForm from "./MyCustomForm";
import { DialogFooter } from "office-ui-fabric-react";
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp } from '@pnp/sp';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import pnp, { sp, Item, ItemAddResult, ItemUpdateResult,Web } from "sp-pnp-js";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { autobind } from 'office-ui-fabric-react';
import { getGUID } from "@pnp/common";
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PrimaryButton } from 'office-ui-fabric-react';
import * as pnp from 'sp-pnp-js';
import { IPnPPeoplePickerProps } from './IPnPPeoplePickerProps';
import { IPnPPeoplePickerState } from './IPnPPeoplePickerState';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export default class MyCustomPanel extends React.Component<IPnPPeoplePickerProps, IPnPPeoplePickerState> {
  private _context = WebPartContext;
  //   public render(): void {
  //     ReactDOM.render(<Panel
  //       isLightDismiss={ true }
  //       isOpen={ true }
  //       type={PanelType.smallFixedFar}
  //       onDismissed={ () => this.close()}>  
  //     {/* <MycustomForm description="abc" /> */}
  //           <h2>New Item</h2>
  //           <Label htmlFor="title" required>Title</Label>
  //           <TextField id="title" ariaLabel="text field" />
  //           <Label htmlFor="assignedto">Assigned To</Label>
  //     </Panel>, this.domElement);
  //   } 
  private title: string = null;
  constructor(props: IPnPPeoplePickerProps, _state: IPnPPeoplePickerState) {
    super(props);
    this.state = {
      firstDayOfWeek: DayOfWeek.Monday,
      addUsers: [],
      saving: false,
      date: undefined,
      title: "",
      dpvalue: "",
      dropvalue: "",
      dudate: new Date(),
    };
  }

  public render(): React.ReactElement<IPnPPeoplePickerProps> {
    let { isOpen, context } = this.props;
    return (
      <Panel isOpen={isOpen} type={PanelType.medium}>
        <h2>New Item</h2>
        <Label htmlFor="title" required>Title</Label>
        <TextField id="title" value={this.state.title} ariaLabel="text field" onChange={this.testField} />
        <Label htmlFor="assignedto">Assigned To</Label>
        <PeoplePicker
          context={this.props.context}
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
          onChange={this._onChange}
          placeholder="Select an option"
          options={[
            { key: 'Active', text: 'Active' },
            { key: 'Resolved', text: 'Resolved' },
            { key: 'Closed', text: 'Closed' },
          ]}
          defaultSelectedKey={this.state.dpvalue}
        // defaultValue  ={this.state.selectedItem}
        // styles={{ dropdown: { width: 300 } }}
        />
        {/* {this.state.dpvalue} */}
        <Label htmlFor="priority">Priority</Label>
        <Dropdown
          id="priority"
          onChange={this.onchange}
          placeholder="Select an option"
          options={[
            { key: '(1)High', text: '(1)High' },
            { key: '(2)Normal', text: '(2)Normal' },
            { key: '(3)Low', text: '(3)Low' },
          ]}
          defaultSelectedKey={this.state.dropvalue}
        />
        <DateTimePicker label="Due Date"
          dateConvention={DateConvention.DateTime}
          timeConvention={TimeConvention.Hours12}
          value={this.state.dudate}
          onChange={this._handleChange}
          isMonthPickerVisible={false}
          timeDisplayControlType={TimeDisplayControlType.Dropdown}
        // onSelectDate={this._onSelectDate}
        />
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={this._onCancel} />
          <PrimaryButton text="save" onClick={this._onsave} />
        </DialogFooter>
      </Panel>);
  }

  @autobind private _onCancel() {
    this.props.onClose();
  }

  @autobind private _onsave() {
    pnp.sp.web.lists.getByTitle('issueTest').items.add({
      Title: this.state.title,
      // AssignedTo : document.getElementById('assignedto')["value"],   
      Status: this.state.dpvalue,
      Priority: this.state.dropvalue,
      DueDate: this.state.dudate,
    });
    alert("Record with Profile Name : " + document.getElementById('title')["value"] + " Added !");
  }
  private _handleChange = (date: Date | null | undefined): void => {
    this.setState({ dudate: date });
  };
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
    this.setState({ addUsers: items });
  }
  private testField = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, issueF?: string): void => {
    this.setState({ title: issueF });

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
  private _onChange = (_event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    this.setState({ dpvalue: item.text });

  };
  private onchange = (_event: React.FormEvent<HTMLDivElement>, Item: IDropdownOption): void => {
    console.log(`Selection change: ${Item.text} ${Item.selected ? 'selected' : 'unselected'}`);
    this.setState({ dropvalue: Item.text });
  };

}
// function _alertClicked(): void {
//   //alert('Clicked');
//   let priority = issuestatus;
//   pnp.sp.web.lists.getByTitle('issueTest').items.add({
//     Title: this.state.title,
//     // AssignedTo : document.getElementById('assignedto')["value"],   
//     Status: this.state.dpvalue,
//     Priority: this.state.dropvalue,
//     // DueDate: document.getElementById('duedate')["value"]
//   });
//   alert("Record with Profile Name : " + document.getElementById('title')["value"] + " Added !");
// }