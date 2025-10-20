declare interface IEmployeeDirectoryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'EmployeeDirectoryWebPartStrings' {
  const strings: IEmployeeDirectoryWebPartStrings;
  export = strings;
}
