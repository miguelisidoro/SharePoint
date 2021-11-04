declare interface IInOfficeSpFxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  IDFieldLabel: string;
  CommandbarEditLabel: string;
  CommandbarViewLabel: string;
  CommandbarDeleteLabel: string;
  DateFieldLabel:string;
  NotesFieldLabel: string;
}

declare module 'InOfficeSpFxWebPartStrings' {
  const strings: IInOfficeSpFxWebPartStrings;
  export = strings;
}
