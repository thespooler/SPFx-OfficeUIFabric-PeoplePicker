declare interface IOfficeUiFabricPeoplePickerStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TypePickerLabel: string;
  principalTypeUserLabel: string;
  principalTypeSharePointGroupLabel: string;
  principalTypeSecurityGroupLabel: string;
  principalTypeDistributionListLabel: string;
  numberOfItemsFieldLabel: string;
  suggestions: string;
  noResults: string;
  loading: string;
}

declare module 'officeUiFabricPeoplePickerStrings' {
  const strings: IOfficeUiFabricPeoplePickerStrings;
  export = strings;
}
