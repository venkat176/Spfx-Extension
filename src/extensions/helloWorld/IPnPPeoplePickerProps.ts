import { WebPartContext } from '@microsoft/sp-webpart-base'; 
import { CompactPeoplePicker, IPersonaProps, IBasePickerSuggestionsProps } from 'office-ui-fabric-react';
import ListViewCommandSetContext from "@microsoft/sp-application-base/lib/extensibility/ApplicationCustomizerContext";
import { ExtensionContext } from '@microsoft/sp-extension-base';


export interface IPnPPeoplePickerProps {  
  description: string;  
  context: WebPartContext; 
  onClose: () => void;
  isOpen: boolean;	    
  currentTitle: string;	 
  maxNrOfUsers:number;   
  cnt : any;
}  