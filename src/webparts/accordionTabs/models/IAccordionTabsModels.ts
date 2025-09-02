/**
 * Enum for view types
 */
export enum ViewType {
  Accordion = "accordion",
  Tabs = "tabs"
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
}

/**
 * Interface for section editor props
 */
export interface ISectionEditorProps {
  section: ISection;
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
}

/**
 * Interface for tabs view props
 */
export interface ITabsViewProps {
  sections: ISection[];
  displayMode: number;
  onSectionsChanged: (sections: ISection[]) => void;
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
