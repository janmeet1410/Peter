export interface IUserProperties {
  map(arg0: (el: any) => JSX.Element): unknown;
  // Department: string;
  PictureUrl: string;
  Title: string;
  DisplayName: string;
  Email: string;
  WorkPhone?: string;
  MobPhone?:string;
  JobTitle?:string;
  OfficeLocation?:string;
  DelveId:string;
  // Location?: string;
  // Manager?:string;
  // Address?:string;
  // ShowInEmpDir?:string;
}
