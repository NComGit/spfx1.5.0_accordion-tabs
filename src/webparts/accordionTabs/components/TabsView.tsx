import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { ContextualMenu, IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ITabsViewProps, ITabsViewState, ISection, TabsDefaultActive } from '../models/IAccordionTabsModels';
import { SectionEditor } from './SectionEditor';
import styles from './TabsView.module.scss';

/**
 * Tabs view component that displays sections in tab format
 */
export class TabsView extends React.Component<ITabsViewProps, ITabsViewState> {

  private _optionsButtonElement: HTMLElement | null = null;
  private _tabHeadersContainerRef: HTMLDivElement | null = null;
  private _isInitialLoad: boolean = true;

  constructor(props: ITabsViewProps) {
    super(props);

    this.state = {
      activeTabIndex: 0,
      editingSection: null,
      showSectionEditor: false,
      showContextualMenu: false,
      contextualMenuTarget: null,
      contextualMenuSection: null,
      contextualMenuIndex: -1,
      isDragging: false,
      dragStartX: 0,
      dragStartY: 0
    };

    // Bind methods
    this.onTabClick = this.onTabClick.bind(this);
    this.onAddSection = this.onAddSection.bind(this);
    this.onEditSection = this.onEditSection.bind(this);
    this.onDeleteSection = this.onDeleteSection.bind(this);
    this.onSectionSave = this.onSectionSave.bind(this);
    this.onSectionCancel = this.onSectionCancel.bind(this);
    this.moveSection = this.moveSection.bind(this);
    this.onShowContextualMenu = this.onShowContextualMenu.bind(this);
    this.onHideContextualMenu = this.onHideContextualMenu.bind(this);
    this.getContextualMenuItems = this.getContextualMenuItems.bind(this);
    this.onMouseDown = this.onMouseDown.bind(this);
    this.onMouseMove = this.onMouseMove.bind(this);
    this.onMouseUp = this.onMouseUp.bind(this);
  }

  public componentDidMount(): void {
    this.initializeActiveTab();
  }

  public componentDidUpdate(prevProps: ITabsViewProps): void {
    // Re-initialize active tab only if configuration changed or sections were added/removed
    // Don't reinitialize for simple reordering (which moveSection handles correctly)
    const sectionsChanged = prevProps.sections !== this.props.sections;
    const configChanged = prevProps.tabsDefaultActive !== this.props.tabsDefaultActive ||
                         prevProps.tabsChosenTab !== this.props.tabsChosenTab;
    
    if (configChanged) {
      // Configuration changed - reinitialize active tab
      this.initializeActiveTab();
    } else if (sectionsChanged) {
      // Check if sections were added/removed (not just reordered)
      const prevSectionIds = prevProps.sections.map(s => s.id).sort();
      const currentSectionIds = this.props.sections.map(s => s.id).sort();
      const sectionIdsChanged = JSON.stringify(prevSectionIds) !== JSON.stringify(currentSectionIds);
      
      if (sectionIdsChanged) {
        // Sections were added or removed - reinitialize active tab
        this.initializeActiveTab();
      } else {
        // Just reordering - ensure active tab is still valid but don't reset to default
        this.ensureValidActiveTab();
      }
    }
  }

  private initializeActiveTab(): void {
    const { sections, tabsDefaultActive, tabsChosenTab } = this.props;
    
    if (sections.length === 0) {
      this.setState((prevState) => ({ ...prevState, activeTabIndex: 0 }));
      return;
    }

    let activeIndex = 0;

    switch (tabsDefaultActive) {
      case TabsDefaultActive.First:
        activeIndex = 0;
        break;
      
      case TabsDefaultActive.Last:
        activeIndex = sections.length - 1;
        break;
      
      case TabsDefaultActive.Chosen:
        // Find the section with the chosen tab ID
        if (tabsChosenTab) {
          // Find the index of the section with the matching ID
          let chosenIndex = -1;
          for (let i = 0; i < sections.length; i++) {
            if (sections[i].id === tabsChosenTab) {
              chosenIndex = i;
              break;
            }
          }
          activeIndex = chosenIndex >= 0 ? chosenIndex : 0; // Fallback to first tab if not found
        } else {
          activeIndex = 0; // Fallback to first tab if no chosen tab specified
        }
        break;
      
      default:
        activeIndex = 0;
        break;
    }

    this.setState((prevState) => ({ ...prevState, activeTabIndex: activeIndex }), () => {
      // Scroll to make the active tab visible after state update
      // Use instant scroll on initial load, smooth animation for subsequent changes
      this.scrollActiveTabIntoView(!this._isInitialLoad);
      this._isInitialLoad = false; // Mark as no longer initial load
    });
  }

  private scrollActiveTabIntoView(useAnimation: boolean = true): void {
    // Use setTimeout to ensure the DOM has been updated
    setTimeout(() => {
      if (this._tabHeadersContainerRef && this.state.activeTabIndex >= 0) {
        const activeTabElement = this._tabHeadersContainerRef.children[this.state.activeTabIndex] as HTMLElement;
        if (activeTabElement) {
          activeTabElement.scrollIntoView({
            behavior: useAnimation ? 'smooth' : 'auto',
            block: 'nearest',
            inline: 'start'
          });
        }
      }
    }, 0);
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

  private onShowContextualMenu(event: React.MouseEvent<HTMLButtonElement>, section: ISection, index: number): void {
    event.preventDefault();
    event.stopPropagation();
    
    // Store the button element reference
    this._optionsButtonElement = event.currentTarget as HTMLElement;
    
    this.setState((prevState) => ({
      ...prevState,
      showContextualMenu: true,
      contextualMenuTarget: this._optionsButtonElement,
      contextualMenuSection: section,
      contextualMenuIndex: index
    }));
  }

  private onHideContextualMenu(): void {
    this.setState((prevState) => ({
      ...prevState,
      showContextualMenu: false,
      contextualMenuTarget: null,
      contextualMenuSection: null,
      contextualMenuIndex: -1
    }));
  }

  private getContextualMenuItems(): IContextualMenuItem[] {
    const { contextualMenuSection, contextualMenuIndex } = this.state;
    const { sections } = this.props;
    
    if (!contextualMenuSection) return [];

    const sortedSections = [...sections].sort((a, b) => a.order - b.order);
    const menuItems: IContextualMenuItem[] = [];

    // Edit option
    menuItems.push({
      key: 'edit',
      name: 'Edit',
      iconProps: { iconName: 'Edit' },
      onClick: () => {
        this.onHideContextualMenu();
        this.onEditSection(contextualMenuSection);
      }
    });

    // Move left option (only show if not the first tab)
    if (contextualMenuIndex > 0) {
      menuItems.push({
        key: 'moveLeft',
        name: 'Move Left',
        iconProps: { iconName: 'Back' },
        onClick: () => {
          this.onHideContextualMenu();
          this.moveSection(contextualMenuSection.id, 'left');
        }
      });
    }

    // Move right option (only show if not the last tab)
    if (contextualMenuIndex < sortedSections.length - 1) {
      menuItems.push({
        key: 'moveRight',
        name: 'Move Right',
        iconProps: { iconName: 'Forward' },
        onClick: () => {
          this.onHideContextualMenu();
          this.moveSection(contextualMenuSection.id, 'right');
        }
      });
    }

    // Delete option
    menuItems.push({
      key: 'delete',
      name: 'Delete',
      iconProps: { iconName: 'Delete' },
      onClick: () => {
        this.onHideContextualMenu();
        this.onDeleteSection(contextualMenuSection.id);
      }
    });

    return menuItems;
  }

  private onMouseDown(event: React.MouseEvent<HTMLElement>): void {
    event.preventDefault();
    
    this.setState((prevState) => ({
      ...prevState,
      isDragging: true,
      dragStartX: event.clientX,
      dragStartY: event.clientY
    }));

    // Add global mouse event listeners
    document.addEventListener('mousemove', this.onMouseMove);
    document.addEventListener('mouseup', this.onMouseUp);
  }

  private onMouseMove(event: MouseEvent): void {
    if (!this.state.isDragging) return;

    // Simple drag feedback - you could enhance this with visual feedback
    const deltaX = event.clientX - this.state.dragStartX;
    const deltaY = event.clientY - this.state.dragStartY;
    
    // For now, we'll use a simple threshold to determine direction
    if (Math.abs(deltaX) > 50) {
      // Determine if we're moving left or right
      const direction = deltaX > 0 ? 'right' : 'left';
      
      if (this.state.contextualMenuSection) {
        this.moveSection(this.state.contextualMenuSection.id, direction);
      }
      
      this.onMouseUp(event);
    }
  }

  private onMouseUp(event: MouseEvent): void {
    this.setState((prevState) => ({
      ...prevState,
      isDragging: false,
      dragStartX: 0,
      dragStartY: 0
    }));

    // Remove global mouse event listeners
    document.removeEventListener('mousemove', this.onMouseMove);
    document.removeEventListener('mouseup', this.onMouseUp);
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
        <div className={styles.tabHeaders} ref={(ref) => { this._tabHeadersContainerRef = ref; }}>
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
                    iconProps={{ iconName: 'More' }}
                    title="Section Options"
                    onClick={(event: any) => this.onShowContextualMenu(event, section, index)}
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
            key={editingSection ? `edit-${editingSection.id}` : 'new-section'}
            section={editingSection}
            isVisible={showSectionEditor}
            onSave={this.onSectionSave}
            onCancel={this.onSectionCancel}
          />
        )}

        {this.state.showContextualMenu && this.state.contextualMenuTarget && (
          <ContextualMenu
            target={this.state.contextualMenuTarget}
            items={this.getContextualMenuItems()}
            onDismiss={this.onHideContextualMenu}
            isBeakVisible={true}
            directionalHint={6} // DirectionalHint.bottomAutoEdge
            gapSpace={0}
            beakWidth={16}
          />
        )}
      </div>
    );
  }
}
