declare interface IOfficeHeroWebPartStrings {
  PropertyPaneDescription: string;
  DataSourceGroupName: string;
  OfficeSelectLabel: string;
  OfficeSelectDescription: string;
  NoOfficesHint: string;
}

declare module 'OfficeHeroWebPartStrings' {
  const strings: IOfficeHeroWebPartStrings;
  export = strings;
}
