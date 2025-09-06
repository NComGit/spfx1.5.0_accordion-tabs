import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IAccordionViewProps, IAccordionViewState, ISection } from '../models/IAccordionTabsModels';
import { SectionEditor } from './SectionEditor';
import styles from './AccordionView.module.scss';

/**
 * Accordion view component that displays sections in collapsible accordion format
 */
export class AccordionView extends React.Component<IAccordionViewProps, IAccordionViewState> {

  constructor(props: IAccordionViewProps) {
    super(props);

    this.state = {
      expandedSections: {},
      editingSection: null,
      showSectionEditor: false
    };

    // Bind methods
    this.toggleSection = this.toggleSection.bind(this);
    this.onAddSection = this.onAddSection.bind(this);
    this.onEditSection = this.onEditSection.bind(this);
    this.onDeleteSection = this.onDeleteSection.bind(this);
    this.onSectionSave = this.onSectionSave.bind(this);
    this.onSectionCancel = this.onSectionCancel.bind(this);
    this.moveSection = this.moveSection.bind(this);
  }

  public componentDidMount(): void {
    // Expand first section by default
    if (this.props.sections.length > 0) {
      const firstSection = this.props.sections[0];
      this.setState((prevState) => ({
        ...prevState,
        expandedSections: { [firstSection.id]: true }
      }));
    }
  }

  private toggleSection(sectionId: string): void {
    this.setState((prevState) => ({
      ...prevState,
      expandedSections: {
        ...prevState.expandedSections,
        [sectionId]: !prevState.expandedSections[sectionId]
      }
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

  private moveSection(sectionId: string, direction: 'up' | 'down'): void {
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
    
    const newIndex = direction === 'up' ? currentIndex - 1 : currentIndex + 1;
    if (newIndex < 0 || newIndex >= sortedSections.length) return;

    // Swap orders
    const temp = sortedSections[currentIndex].order;
    sortedSections[currentIndex].order = sortedSections[newIndex].order;
    sortedSections[newIndex].order = temp;

    this.props.onSectionsChanged(sortedSections);
  }

  private renderSectionContent(content: string): React.ReactElement<any> {
    return (
      <div 
        className={styles.sectionContent}
        dangerouslySetInnerHTML={{ __html: content }}
      />
    );
  }

  public render(): React.ReactElement<IAccordionViewProps> {
    const { sections, displayMode } = this.props;
    const { expandedSections, editingSection, showSectionEditor } = this.state;
    const isEditMode = displayMode === DisplayMode.Edit;

    // Sort sections by order
    const sortedSections = [...sections].sort((a, b) => a.order - b.order);

    if (sections.length === 0 && !isEditMode) {
      return (
        <div className={styles.emptyState}>
          <Icon iconName="DocumentSearch" className={styles.emptyIcon} />
          <div className={styles.emptyMessage}>No sections configured</div>
        </div>
      );
    }

    return (
      <div className={styles.accordionContainer}>
        {sortedSections.map((section, index) => (
          <div key={section.id} className={styles.accordionSection}>
            <div 
              className={`${styles.accordionHeader} ${expandedSections[section.id] ? styles.expanded : ''}`}
              onClick={() => this.toggleSection(section.id)}
            >
              <div className={styles.sectionTitle}>
                <Icon 
                  iconName={expandedSections[section.id] ? "ChevronDown" : "ChevronRight"} 
                  className={styles.chevronIcon}
                />
                <span>{section.title}</span>
              </div>
              
              {isEditMode && (
                <div className={styles.sectionActions} onClick={(e) => e.stopPropagation()}>
                  <IconButton
                    iconProps={{ iconName: 'Edit' }}
                    title="Edit Section"
                    onClick={() => this.onEditSection(section)}
                    className={styles.actionButton}
                  />
                  {index > 0 && (
                    <IconButton
                      iconProps={{ iconName: 'Up' }}
                      title="Move Up"
                      onClick={() => this.moveSection(section.id, 'up')}
                      className={styles.actionButton}
                    />
                  )}
                  {index < sortedSections.length - 1 && (
                    <IconButton
                      iconProps={{ iconName: 'Down' }}
                      title="Move Down"
                      onClick={() => this.moveSection(section.id, 'down')}
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

            {expandedSections[section.id] && (
              <div className={styles.accordionBody}>
                {section.content ? 
                  this.renderSectionContent(section.content) :
                  <div className={styles.emptyContent}>No content available</div>
                }
              </div>
            )}
          </div>
        ))}

        {isEditMode && (
          <div className={styles.addSectionContainer}>
            <IconButton
              iconProps={{ iconName: 'Add' }}
              text="Add Section"
              onClick={this.onAddSection}
              className={styles.addSectionButton}
            />
          </div>
        )}

        {showSectionEditor && (
          <SectionEditor
            key={editingSection ? `edit-${editingSection.id}` : 'new-section'}
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
