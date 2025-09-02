declare interface IAccordionTabsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  
  // View Configuration
  ViewConfigurationGroupName: string;
  ViewTypeLabel: string;
  ViewTypeAccordion: string;
  ViewTypeTabs: string;
  
  // Sections Management
  SectionsGroupName: string;
  AddSectionButton: string;
  SectionTitleLabel: string;
  EditSectionButton: string;
  DeleteSectionButton: string;
  MoveUpButton: string;
  MoveDownButton: string;
  DefaultSectionTitle: string;
  
  // Validation Messages
  SectionTitleRequiredError: string;
  SectionTitleLengthError: string;
}

declare module 'AccordionTabsWebPartStrings' {
  const strings: IAccordionTabsWebPartStrings;
  export = strings;
}
