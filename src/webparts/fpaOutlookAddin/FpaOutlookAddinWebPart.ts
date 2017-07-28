import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as pnp from "sp-pnp-js";
import { IFPADropdownData } from "./DataModel";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IWebPartContext,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import office from 'office-js';
import * as strings from 'fpaOutlookAddinStrings';
import FpaOutlookAddin from './components/FpaOutlookAddin';
import { IFpaOutlookAddinProps } from './components/IFpaOutlookAddinProps';
import { IFpaOutlookAddinWebPartProps } from './IFpaOutlookAddinWebPartProps';

export default class FpaOutlookAddinWebPart extends BaseClientSideWebPart<IFpaOutlookAddinWebPartProps> {
  private from: string;
  private attachments: Array<any>;
  private body: string;
  private office: office;
  private subject: string;
  public onInit<T>(): Promise<T> {
    return new Promise<T>((resolve: (args: T) => void, reject: (error: Error) => void) => {
      window["Office"]["initialize"] = () => {
        debugger;
        this.office = window["Office"];
        this.attachments = window["Office"].context.mailbox.item.attachments;
        this.from = window["Office"].context.mailbox.item.from.emailAddress;
        this.subject = window["Office"].context.mailbox.item.subject;
        window["Office"].context.mailbox.item.body.getAsync(
          "html",
          { asyncContext: "This is passed to the callback" },
          (result) => {
            debugger;
            this.body = result.value;
            resolve(window["Office"]);// or undefined
          }

        );
      };
    });

  }
  public getDropdownData(): Promise<IFPADropdownData> {
    let region: Array<string>;
    let contactLocation: Array<string>;
    let states: Array<string>;
    let countries: Array<string>;
    let affectedProducts: Array<string>;
    let applicationTypes: Array<string>;
    let batch = pnp.sp.createBatch();
    pnp.sp.web.lists.getByTitle("Applications").items.inBatch(batch).get().then((items) => {
      debugger;
      applicationTypes = items.map((item) => {
        return item.Title
      })
    }).catch((error) => {
      console.log("error fetching applictaions");
      console.log(error);
    });
    pnp.sp.web.lists.getByTitle("Countries").items.inBatch(batch).get().then((items) => {
      debugger;
      countries = items.map((item) => {
        return item.Title
      })
    }).catch((error) => {
      console.log("error fetching Countries");
      console.log(error);
    });
    pnp.sp.web.lists.getByTitle("Products").items.inBatch(batch).get().then((items) => {
      debugger;
      affectedProducts = items.map((item) => {
        return item.Title
      })
    }).catch((error) => {
      console.log("error fetching Products");
      console.log(error);
    });
    pnp.sp.web.lists.getByTitle("States").items.inBatch(batch).get().then((items) => {
      debugger;
      states = items.map((item) => {
        return item.Title
      })
    }).catch((error) => {
      console.log("error fetching States");
      console.log(error);
    });

    pnp.sp.web.lists.getByTitle("Request Regions").items.inBatch(batch).get().then((items) => {
      debugger;
      region = items.map((item) => {
        return item.Title
      })
    }).catch((error) => {
      console.log("error fetching Request Regions");
      console.log(error);
    });
    pnp.sp.web.lists.getByTitle("Contact Locations").items.inBatch(batch).get().then((items) => {
      debugger;
      contactLocation = items.map((item) => {
        return item.Title
      })
    }).catch((error) => {
      console.log("error fetching ontact Locations");
      console.log(error);
    });

    return batch.execute().then(() => {
      let dropdownData: IFPADropdownData = {
        state: states,
        region: region,
        applicationTypes: applicationTypes,
        affectedProducts: affectedProducts,
        contactLocation: contactLocation,
        country: countries

      }

      return dropdownData;


    });



  }
  public render(): void {
    debugger;
    const props: IFpaOutlookAddinProps = {
      from: this.from,
      body: this.body,
      attachments: this.attachments,
      office: this.office,
      subject: this.subject,
      getDropdownData: this.getDropdownData.bind(this),
      save: null

    }
    const element: React.ReactElement<IFpaOutlookAddinProps> = React.createElement(
      FpaOutlookAddin, props
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
