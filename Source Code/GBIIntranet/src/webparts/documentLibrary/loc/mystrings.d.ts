declare interface IDocumentLibraryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'DocumentLibraryWebPartStrings' {
  const strings: IDocumentLibraryWebPartStrings;
  export = strings;
}
