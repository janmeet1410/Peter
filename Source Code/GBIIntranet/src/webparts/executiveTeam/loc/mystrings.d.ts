declare interface IExecutiveTeamWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ExecutiveTeamWebPartStrings' {
  const strings: IExecutiveTeamWebPartStrings;
  export = strings;
}
