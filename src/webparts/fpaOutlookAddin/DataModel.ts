export interface IFPADropdownData {
    region: Array<ILookupField>;
    contactLocation: Array<ILookupField>;
    state: Array<ILookupField>;
    country: Array<ILookupField>;
    affectedProducts: Array<ILookupField>;
    applicationTypes: Array<ILookupField>;
}
export interface IFPAFields {
    region?: ILookupField;
    contactLocation?: ILookupField;
    contact_FName?: string;
    contact_LName?: string;
    contact_Address?: string;
    contact_City?: string;
    contact_State?: ILookupField;
    contact_Country?: ILookupField;
    contact_PostalCode?: string;
    contact_Phone?: string;
    contact_Fax?: string;
    contactEmailAddress?: string;

}
export interface ILookupField {
    key: number;
    text: string;
}