import office from 'office-js';
import {IFPADropdownData} from "../DataModel";
export interface IFpaOutlookAddinProps {
   from: string;
   attachments: Array<any>;
   body: string;
   office: office;
   subject:string;
   getDropdownData:()=>Promise<IFPADropdownData>;
   save:()=>Promise<any>;
}
