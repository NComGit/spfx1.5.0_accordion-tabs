import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ITabsViewProps, ISection } from '../models/IAccordionTabsModels';
import { SectionEditor } from './SectionEditor';
import styles from './TabsView.module.scss';

interface ITabsViewState {
  activeTabIndex: number;
  editingSection: ISection | null;
  showSectionEditor: boolean;
}

/**
 * Tabs view component that displays sections in tab format
 */
export class TabsView extends React.Component<ITabsViewProps, ITabsViewState> {

  constructor(props: ITabsViewProps) {
    super(props);

    this.state = {
      activeTabIndex: 0,
      editingSection: null,
      showSectionEditor: false
    };

    // Bind methods
    this.onTabClick = this.onTabClick.bind(this);
    this.onAddSection = this.onAddSection.bind(this);
    this.onEditSection = this.onEditSection.bind(this);
    this.onDeleteSection = this.onDeleteSection.bind(this);
    this.onSectionSave = this.onSectionSave.bind(this);
    this.onSectionCancel = this.onSectionCancel.bind(this);
    this.moveSection = this.moveSection.bind(this);
  }

  public componentDidMount(): void {
    // Ensure active tab is within bounds
    this.ensureValidActiveTab();
  }

  public componentDidUpdate(prevProps: ITabsViewProps): void {
    // Update active tab if sections changed
    if (prevProps.sections.length !== this.props.sections.length) {
      this.ensureValidActiveTab();
    }
  }

  private ensureValidActiveTab(): void {
    const { sections } = this.props;
    const { activeTabIndex } = this.state;
    
    if (sections.length === 0) {
      this.setState((prevState) => ({ ...prevState, activeTabIndex: 0 }));
    } else if (activeTabIndex >= sections.length) {
      this.setState((prevState) => ({ ...prevState, activeTabIndex: sections.length - 1 }));
    }
  }

  private onTabClick(index: number): void {
    this.setState((prevState) => ({
      ...prevState,
      activeTabIndex: index
    }));
  }

  private onAddSection(): void {
    this.setState((prevState) => ({
      ...prevState,
      editingSection: null,
      showSectionEditor: true
    }));
  }

  private onEditSection(section: ISection): void {
    this.setState((prevState) => ({
      ...prevState,
      editingSection: section,
      showSectionEditor: true
    }));
  }

  private onDeleteSection(sectionId: string): void {
    const updatedSections = this.props.sections.filter(s => s.id !== sectionId);
    this.props.onSectionsChanged(updatedSections);
  }

  private onSectionSave(section: ISection): void {
    const { sections } = this.props;
    let updatedSections: ISection[];

    if (this.state.editingSection) {
      // Update existing section
      updatedSections = sections.map(s => s.id === section.id ? section : s);
    } else {
      // Add new section
      const maxOrder = sections.length > 0 ? Math.max(...sections.map(s => s.order)) : -1;
      section.order = maxOrder + 1;
      updatedSections = [...sections, section];
      
      // Set new tab as active
      this.setState((prevState) => ({
        ...prevState,
        activeTabIndex: updatedSections.length - 1
      }));
    }

    this.props.onSectionsChanged(updatedSections);
    this.setState((prevState) => ({
      ...prevState,
      editingSection: null,
      showSectionEditor: false
    }));
  }

  private onSectionCancel(): void {
    this.setState((prevState) => ({
      ...prevState,
      editingSection: null,
      showSectionEditor: false
    }));
  }

  private moveSection(sectionId: string, direction: 'left' | 'right'): void {
    const { sections } = this.props;
    const sortedSections = [...sections].sort((a, b) => a.order - b.order);
    
    // Find current index manually for TypeScript 2.4.2 compatibility
    let currentIndex = -1;
    for (let i = 0; i < sortedSections.length; i++) {
      if (sortedSections[i].id === sectionId) {
        currentIndex = i;
        break;
      }
    }
    
    if (currentIndex === -1) return;
    
    const newIndex = direction === 'left' ? currentIndex - 1 : currentIndex + 1;
    if (newIndex < 0 || newIndex >= sortedSections.length) return;

    // Swap orders
    const temp = sortedSections[currentIndex].order;
    sortedSections[currentIndex].order = sortedSections[newIndex].order;
    sortedSections[newIndex].order = temp;

    // Update active tab index to follow the moved tab
    if (this.state.activeTabIndex === currentIndex) {
      this.setState((prevState) => ({ ...prevState, activeTabIndex: newIndex }));
    } else if (this.state.activeTabIndex === newIndex) {
      this.setState((prevState) => ({ ...prevState, activeTabIndex: currentIndex }));
    }

    this.props.onSectionsChanged(sortedSections);
  }

  private renderSectionContent(content: string): React.ReactElement<any> {
    return (
      <div 
        className={styles.tabContent}
        dangerouslySetInnerHTML={{ __html: content }}
      />
    );
  }

  public render(): React.ReactElement<ITabsViewProps> {
    const { sections, displayMode } = this.props;
    const { activeTabIndex, editingSection, showSectionEditor } = this.state;
    const isEditMode = displayMode === DisplayMode.Edit;

    // Sort sections by order
    const sortedSections = [...sections].sort((a, b) => a.order - b.order);
    const activeSection = sortedSections[activeTabIndex];

    if (sections.length === 0 && !isEditMode) {
      return (
        <div className={styles.emptyState}>
          <Icon iconName="DocumentSearch" className={styles.emptyIcon} />
          <div className={styles.emptyMessage}>No sections configured</div>
        </div>
      );
    }

    return (
      <div className={styles.tabsContainer}>
        {/* Tab Headers */}
        <div className={styles.tabHeaders}>
          {sortedSections.map((section, index) => (
            <div 
              key={section.id}
              className={`${styles.tabHeader} ${index === activeTabIndex ? styles.active : ''}`}
              onClick={() => this.onTabClick(index)}
            >
              <span className={styles.tabTitle}>{section.title}</span>
              
              {isEditMode && (
                <div className={styles.tabActions} onClick={(e) => e.stopPropagation()}>
                  <IconButton
                    iconProps={{ iconName: 'Edit' }}
                    title="Edit Section"
                    onClick={() => this.onEditSection(section)}
                    className={styles.actionButton}
                  />
                  {index > 0 && (
                    <IconButton
                      iconProps={{ iconName: 'Back' }}
                      title="Move Left"
                      onClick={() => this.moveSection(section.id, 'left')}
                      className={styles.actionButton}
                    />
                  )}
                  {index < sortedSections.length - 1 && (
                    <IconButton
                      iconProps={{ iconName: 'Forward' }}
                      title="Move Right"
                      onClick={() => this.moveSection(section.id, 'right')}
                      className={styles.actionButton}
                    />
                  )}
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    title="Delete Section"
                    onClick={() => this.onDeleteSection(section.id)}
                    className={styles.actionButton}
                  />
                </div>
              )}
            </div>
          ))}
          
          {isEditMode && (
            <div className={styles.addTabContainer}>
              <IconButton
                iconProps={{ iconName: 'Add' }}
                title="Add Section"
                onClick={this.onAddSection}
                className={styles.addTabButton}
              />
            </div>
          )}
        </div>

        {/* Tab Content */}
        <div className={styles.tabContentContainer}>
          {activeSection ? (
            <div className={styles.tabPane}>
              {activeSection.content ? 
                this.renderSectionContent(activeSection.content) :
                <div className={styles.emptyContent}>No content available</div>
              }
            </div>
          ) : (
            <div className={styles.emptyContent}>
              {isEditMode ? 'Add a section to get started' : 'No content available'}
            </div>
          )}
        </div>

        {showSectionEditor && (
          <SectionEditor
            section={editingSection}
            isVisible={showSectionEditor}
            onSave={this.onSectionSave}
            onCancel={this.onSectionCancel}
          />
        )}
      </div>
    );
  }
}
