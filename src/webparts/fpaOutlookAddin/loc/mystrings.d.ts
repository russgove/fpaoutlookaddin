declare interface IFpaOutlookAddinStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'fpaOutlookAddinStrings' {
  const strings: IFpaOutlookAddinStrings;
  export = strings;
}
