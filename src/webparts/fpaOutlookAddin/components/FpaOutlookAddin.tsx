import * as React from 'react';
import styles from './FpaOutlookAddin.module.scss';
import { IFpaOutlookAddinProps } from './IFpaOutlookAddinProps';
import { IFpaOutlookAddinState } from './IFpaOutlookAddinState';
import { IFPADropdownData, ILookupField } from "../DataModel";
import { escape } from '@microsoft/sp-lodash-subset';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { Dropdown, IDropdownProps, } from 'office-ui-fabric-react/lib/Dropdown';

export default class FpaOutlookAddin extends React.Component<IFpaOutlookAddinProps, IFpaOutlookAddinState> {
  public componentDidMount() {
    this.props.getDropdownData().then((dropdownData: IFPADropdownData) => {
      this.setState({ dropdownData: dropdownData, fpafields: {} })
    }).catch((err) => {
      debugger;
    })

  }


  public render(): React.ReactElement<IFpaOutlookAddinProps> {
    debugger;
    if (!(this.state && this.state.dropdownData)) {
      return (
        <div>
          Loading...
          </div>
      )
    }

    return (
      <div>
        <table>


          <tr>
            <td>
              <Label>
                Region
                </Label>
            </td>
          </tr>
          <tr>
            <td>

              <Dropdown label=''
                options={this.state.dropdownData.region.map((region) => { return { key: region.key, text: region.text } })}
                onChanged={e => {
                  debugger;
                  this.state.fpafields.region = e as ILookupField;
                  this.setState(this.state);
                }} />

            </td>
          </tr>

          <tr>
            <td>
              <Label>
                Contact Location
                </Label>
            </td>
          </tr>
          <tr>
            <td>

              <Dropdown label=''
                options={this.state.dropdownData.contactLocation.map((contactLocation) => { return { key: contactLocation.key, text: contactLocation.text } })}
                onChanged={e => {
                  debugger;
                  this.state.fpafields.contactLocation = e as ILookupField;
                  this.setState(this.state);
                }} />

            </td>
          </tr>


          <tr>
            <td>
              <Label>
                Contact First Name
                </Label>
            </td>
          </tr>
          <tr>
            <td>

              <TextField value={this.state.fpafields.contact_FName} onChanged={e => {
                this.state.fpafields.contact_FName = e; this.setState(this.state);
              }} />
            </td>
          </tr>



          <tr>
            <td>
              <Label>
                Contact Last Name
                </Label>
            </td>
          </tr>
          <tr>
            <td>

              <TextField value={this.state.fpafields.contact_LName} onChanged={e => {
                this.state.fpafields.contact_LName = e; this.setState(this.state);
              }} />
            </td>
          </tr>



          <tr>
            <td>
              <Label>
                Contact Address
                </Label>
            </td>
          </tr>
          <tr>
            <td>

              <TextField multiline={true} value={this.state.fpafields.contact_Address} onChanged={e => {
                this.state.fpafields.contact_Address = e; this.setState(this.state);
              }} />
            </td>
          </tr>




        </table>
        Subject :{this.props.subject}
        From :{this.props.from}
        Body: {this.props.body}
      </div>
    );
  }
}
