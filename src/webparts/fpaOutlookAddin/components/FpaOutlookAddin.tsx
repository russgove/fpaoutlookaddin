import * as React from 'react';
import styles from './FpaOutlookAddin.module.scss';
import { IFpaOutlookAddinProps } from './IFpaOutlookAddinProps';
import { IFpaOutlookAddinState } from './IFpaOutlookAddinState';
import {IFPADropdownData} from "../DataModel";
import { escape } from '@microsoft/sp-lodash-subset';

export default class FpaOutlookAddin extends React.Component<IFpaOutlookAddinProps, IFpaOutlookAddinState> {
public componentDidMount(){
  this.props.getDropdownData().then((dropdownData:IFPADropdownData)=>{
    this.setState({dropdownData:dropdownData})

  }).catch((err)=>{

  })

}

  public render(): React.ReactElement<IFpaOutlookAddinProps> {
    return (
      <div>
        Subject :{this.props.subject}
        From :{this.props.from}
        Body: {this.props.body}
        
        </div>
    );
  }
}
