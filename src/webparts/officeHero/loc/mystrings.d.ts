declare interface IOfficeHeroWebPartStrings {
  PropertyPaneDescription: string;

  OfficeGroupName: string;
  OfficeFactsGroupName: string;
  FacilitiesNoticeGroupName: string;
  KeyContactsGroupName: string;

  OfficeNameLabel: string;
  AddressLineLabel: string;
  OpeningHoursLabel: string;
  PostRoomHoursLabel: string;
  BackgroundImageLabel: string;
  ChooseImageButton: string;
  ImageAltTextLabel: string;
  ImageAltTextDescription: string;

  FactsLabel: string;
  FactsPanelHeader: string;
  FactsPanelDescription: string;
  FactsManageButton: string;
  FactsIconColumn: string;
  FactsTextColumn: string;

  ShowNoticeLabel: string;
  NoticeLabelLabel: string;
  NoticeTitleLabel: string;
  NoticeDetailLabel: string;
  NoticeCtaTextLabel: string;
  NoticeCtaUrlLabel: string;

  ContactRoleLabelLabel: string;
  ContactPersonLabel: string;
  ContactJobTitleLabel: string;
  Contact1Header: string;
  Contact2Header: string;
  Contact3Header: string;
  Contact4Header: string;

  FieldRequired: string;
  FieldInvalidUrl: string;
}

declare module 'OfficeHeroWebPartStrings' {
  const strings: IOfficeHeroWebPartStrings;
  export = strings;
}
