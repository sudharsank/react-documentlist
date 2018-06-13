declare interface IDocumentListWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DocLibraryFieldLabel: string;
  DocLibraryFieldCalloutContent: string;
  LayoutTypeFieldLabel: string;
  LayoutTypeFieldCalloutContent: string;
  DateFormatFieldLabel: string;
  DateFormatFieldCalloutContent: string;
  ShowFoldersFieldLabel: string;
  ShowFoldersFieldCalloutContent: string;
}

declare module 'DocumentListWebPartStrings' {
  const strings: IDocumentListWebPartStrings;
  export = strings;
}
