import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import IFrameDialog from './IFrameDialogContent';
import * as jquery from 'jquery';
import * as strings from 'WarNewCommandSetCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IWarNewCommandSetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

let listDisplayName = "WAR";
let listName;
const LOG_SOURCE: string = 'WarNewCommandSetCommandSet';

export default class WarNewCommandSetCommandSet extends BaseListViewCommandSet<IWarNewCommandSetCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    listName = this.context.pageContext.list.title;
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_2');

    console.log("Inside onInit : "+listName);
      
    if (listDisplayName.indexOf(listName)>-1) {
      compareOneCommand.visible = true;

      jquery("span.CommandBarItem-commandText").filter(function() { return (jquery(this).text() === 'New') }).parent().hide();
       jquery("span:contains('Export to Excel')" ).parent().hide();
       jquery("span:contains('Quick edit')" ).parent().hide();
       jquery("span:contains('Upload')" ).parent().hide();
       jquery("span:contains('Sync')" ).parent().hide();
    }
    Log.info(LOG_SOURCE, 'Initialized WarNewCommandSetCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_2');
    listName = this.context.pageContext.list.title;

    if (listDisplayName.indexOf(listName)>-1) {

      jquery("span.CommandBarItem-commandText").filter(function() { return (jquery(this).text() === 'New') }).parent().hide();
      jquery("span:contains('Export to Excel')" ).parent().hide();
      jquery("span:contains('Quick edit')" ).parent().hide();
      jquery("span:contains('Upload')" ).parent().hide();
      jquery("span:contains('Sync')" ).parent().hide();
      
    }
    if (compareOneCommand) {
      // This command should be hidden unless no row is selected.
      compareOneCommand.visible = ( listName === listDisplayName && event.selectedRows.length === 0);  //event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_2':
        new IFrameDialog(this.context.pageContext.site.serverRelativeUrl+"/_layouts/15/listform.aspx?PageType=8&ListId="+this.context.pageContext.list.id+"&RootFolder=&IsDlg=1").show();
        //Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
