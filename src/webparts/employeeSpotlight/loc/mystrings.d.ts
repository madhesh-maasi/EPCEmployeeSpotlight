declare interface IEmployeeSpotlightWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  propertyPaneHeading: string;
  selectSiteLableMessage: "Site";
  selectListLableMessage: string;
  employeeEmailcolumnLableMessage: string;
  descriptioncolumnLableMessage: string;
  expirationDateColumnLableMessage: string;
  effectsGroupName: string;
  spotlightBGColorLableMessage:string;
  spotlightFontColorLableMessage:string;
  enableAutoSlideLableMessage: string;
  carouselSpeedLableMessage: string;
}

declare module 'EmployeeSpotlightWebPartStrings' {
  const strings: IEmployeeSpotlightWebPartStrings;
  export = strings;
  propertyPaneHeading : string
}
