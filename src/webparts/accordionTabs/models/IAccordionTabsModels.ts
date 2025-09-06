/**
 * Enum for view types
 */
export enum ViewType {
  Accordion = "accordion",
  Tabs = "tabs"
}

/**
 * Enum for accordion default expanded options
 */
export enum AccordionDefaultExpanded {
  None = "none",
  First = "first", 
  All = "all",
  Chosen = "chosen"
}

/**
 * Enum for tabs default active options
 */
export enum TabsDefaultActive {
  First = "first",
  Last = "last",
  Chosen = "chosen"
}

/**
 * Interface for individual section data
 */
export interface ISection {
  id: string;
  title: string;
  content: string;
  order: number;
}

/**
 * Interface for web part properties
 */
export interface IAccordionTabsWebPartProps {
  viewType: ViewType;
  sections: ISection[];
  allowEdit: boolean;
  // Accordion settings
  accordionDefaultExpanded: AccordionDefaultExpanded;
  accordionChosenSection: string; // ID of chosen section to expand
  // Tabs settings
  tabsDefaultActive: TabsDefaultActive; // Type of default active tab
  tabsChosenTab: string; // ID of chosen tab when "Chosen" is selected
}

/**
 * Interface for the main component props
 */
export interface IAccordionTabsProps {
  viewType: ViewType;
  sections: ISection[];
  displayMode: number; // SPFx DisplayMode
  onSectionsChanged: (sections: ISection[]) => void;
  onConfigureClick: () => void;
  // Accordion settings
  accordionDefaultExpanded: AccordionDefaultExpanded;
  accordionChosenSection: string;
  // Tabs settings
  tabsDefaultActive: TabsDefaultActive;
  tabsChosenTab: string;
}

/**
 * Interface for section editor props
 */
export interface ISectionEditorProps {
  section: ISection | null;
  isVisible: boolean;
  onSave: (section: ISection) => void;
  onCancel: () => void;
}

/**
 * Interface for accordion view props
 */
export interface IAccordionViewProps {
  sections: ISection[];
  displayMode: number;
  onSectionsChanged: (sections: ISection[]) => void;
  accordionDefaultExpanded: AccordionDefaultExpanded;
  accordionChosenSection: string;
}

/**
 * Interface for tabs view props
 */
export interface ITabsViewProps {
  sections: ISection[];
  displayMode: number;
  onSectionsChanged: (sections: ISection[]) => void;
  tabsDefaultActive: TabsDefaultActive;
  tabsChosenTab: string;
}

/**
 * Interface for TinyMCE editor props
 */
export interface ITinyMCEEditorProps {
  value: string;
  onEditorChange: (content: string) => void;
  height?: number;
}

/**
 * Interface for section manager state
 */
export interface ISectionManagerState {
  editingSection: ISection | null;
  showEditor: boolean;
  isLoading: boolean;
}

/**
 * Interface for accordion/tabs component state
 */
export interface IAccordionTabsState {
  sections: ISection[];
  activeTabIndex: number;
  activeAccordionIndex: number;
  editingSection: ISection | null;
  showSectionEditor: boolean;
  isLoading: boolean;
  error: string | null;
}

/**
 * Interface for AccordionTabsComponent state
 */
export interface IAccordionTabsComponentState {
  isLoading: boolean;
  error: string | null;
  sections: ISection[];
}

/**
 * Interface for AccordionView state
 */
export interface IAccordionViewState {
  expandedSections: { [key: string]: boolean };
  editingSection: ISection | null;
  showSectionEditor: boolean;
  showDeleteConfirmation: boolean;
  sectionToDelete: ISection | null;
}

/**
 * Interface for TabsView state
 */
export interface ITabsViewState {
  activeTabIndex: number;
  editingSection: ISection | null;
  showSectionEditor: boolean;
  showContextualMenu: boolean;
  contextualMenuTarget: Element | null;
  contextualMenuSection: ISection | null;
  contextualMenuIndex: number;
  isDragging: boolean;
  dragStartX: number;
  dragStartY: number;
  showDeleteConfirmation: boolean;
  sectionToDelete: ISection | null;
}

/**
 * Interface for SectionEditor state
 */
export interface ISectionEditorState {
  title: string;
  content: string;
  isLoading: boolean;
  hasChanges: boolean;
}
