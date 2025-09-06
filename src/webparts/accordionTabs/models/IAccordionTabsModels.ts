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
  accordionChosenSection: number; // Index of chosen section to expand
  // Tabs settings
  tabsDefaultActive: number; // Index of default active tab
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
  accordionChosenSection: number;
  // Tabs settings
  tabsDefaultActive: number;
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
  accordionChosenSection: number;
}

/**
 * Interface for tabs view props
 */
export interface ITabsViewProps {
  sections: ISection[];
  displayMode: number;
  onSectionsChanged: (sections: ISection[]) => void;
  tabsDefaultActive: number;
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
