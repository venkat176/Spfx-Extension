import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import MyCustomPanel from "./MyCustomPanel";
import { WebPartContext } from '@microsoft/sp-webpart-base';
// import MycustomForm from './MyCustomForm';
import { IPnPPeoplePickerProps } from './IPnPPeoplePickerProps';
import { SPHttpClient, HttpClientResponse } from "@microsoft/sp-http";
// import { sp,Web } from 'sp-pnp-js';
import { sp } from "@pnp/sp";
import { autobind, assign } from '@uifabric/utilities';
import { IPnPPeoplePickerState } from './IPnPPeoplePickerState';
import { ExtensionContext } from '@microsoft/sp-extension-base';
import * as pnp from 'sp-pnp-js';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */

export interface IHelloWorldCommandSetProperties {
  // This is an example; replace with your own properties
  description: string;
  context: WebPartContext;
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'HelloWorldCommandSet';

export default class HelloWorldCommandSet extends BaseListViewCommandSet<IHelloWorldCommandSetProperties> {

  //  protected readonly domElement: HTMLElement;

  // public render(): void {
  //   const element: React.ReactElement<IPnPPeoplePickerProps> = React.createElement(
  //     MycustomForm,
  //     {
  //       description: this.properties.description,
  //       context: this.context
  //     }
  //   );
  //   ReactDOM.render(element, this.domElement);
  // }

  private panelPlaceHolder: HTMLDivElement = null;
  domElement: Element;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized HelloWorldCommandSet');

    console.log(this.context.pageContext.web.absoluteUrl);

    pnp.setup({
      spfxContext: this.context
    });

    // sp.setup({
    //   spfxContext: this.context
    // });    
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));

    return Promise.resolve();
  }

  private _showPanel(itemId: number, currentTitle: string) {
    this._renderPanelComponent({
      isOpen: true, currentTitle, itemId, listId: this.context.pageContext.list.id.toString(), onClose: this._dismissPanel
    });
  }
  @autobind private _dismissPanel() {
    this._renderPanelComponent({ isOpen: false });
  }
  private _renderPanelComponent(props: any) {
    const element: React.ReactElement<IPnPPeoplePickerProps> = React.createElement(MyCustomPanel, assign({
      onClose: null, currentTitle: null, itemId: null, isOpen: false, listId: null
    }, props));
    ReactDOM.render(element, this.panelPlaceHolder);
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        // const dialog: MyCustomPanel = new MyCustomPanel();
        // dialog.show();
        let selectedItem = event.selectedRows[0];
        const listItemId = selectedItem.getValueByName('ID') as number;
        const title = selectedItem.getValueByName("Title");
        this._showPanel(listItemId, title);
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
