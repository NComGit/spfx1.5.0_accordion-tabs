import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IAccordionTabsProps, IAccordionTabsComponentState, ViewType, ISection } from '../models/IAccordionTabsModels';
import { AccordionView } from './AccordionView';
import { TabsView } from './TabsView';
import styles from './AccordionTabsComponent.module.scss';

/**
 * Main component that renders either accordion or tabs view based on configuration
 */
export class AccordionTabsComponent extends React.Component<IAccordionTabsProps, IAccordionTabsComponentState> {

  constructor(props: IAccordionTabsProps) {
    super(props);

    this.state = {
      isLoading: false,
      error: null,
      sections: props.sections || []
    };

    // Bind methods
    this.onSectionsChanged = this.onSectionsChanged.bind(this);
    this.onConfigureClick = this.onConfigureClick.bind(this);
  }

  public componentDidUpdate(prevProps: IAccordionTabsProps): void {
    // Update sections if props changed
    if (prevProps.sections !== this.props.sections) {
      this.setState((prevState) => ({
        ...prevState,
        sections: this.props.sections || []
      }));
    }
  }

  private onSectionsChanged(sections: ISection[]): void {
    // Update local state
    this.setState((prevState) => ({
      ...prevState,
      sections: sections
    }));

    // Notify parent component (web part) of changes
    this.props.onSectionsChanged(sections);
  }

  private onConfigureClick(): void {
    // Notify parent to open property pane
    this.props.onConfigureClick();
  }

  private renderEmptyState(): React.ReactElement<any> {
    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyStateContent}>
          <div className={styles.emptyIcon}>📄</div>
          <h3 className={styles.emptyTitle}>No sections configured</h3>
          <p className={styles.emptyDescription}>
            Add sections to display content in {this.props.viewType === ViewType.Accordion ? 'accordion' : 'tabs'} format.
          </p>
          <button 
            className={styles.configureButton} 
            onClick={this.onConfigureClick}
          >
            Configure Web Part
          </button>
        </div>
      </div>
    );
  }

  private renderErrorState(): React.ReactElement<any> {
    return (
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={true}
        onDismiss={() => this.setState((prevState) => ({ ...prevState, error: null }))}
      >
        <strong>Error:</strong> {this.state.error}
      </MessageBar>
    );
  }

  private renderLoadingState(): React.ReactElement<any> {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Loading content..." />
      </div>
    );
  }

  private renderContent(): React.ReactElement<any> {
    const { viewType, displayMode, accordionDefaultExpanded, accordionChosenSection, tabsDefaultActive } = this.props;
    const { sections } = this.state;

    switch (viewType) {
      case ViewType.Accordion:
        return (
          <AccordionView 
            sections={sections}
            displayMode={displayMode}
            onSectionsChanged={this.onSectionsChanged}
            accordionDefaultExpanded={accordionDefaultExpanded}
            accordionChosenSection={accordionChosenSection}
          />
        );
      
      case ViewType.Tabs:
        return (
          <TabsView 
            sections={sections}
            displayMode={displayMode}
            onSectionsChanged={this.onSectionsChanged}
            tabsDefaultActive={tabsDefaultActive}
          />
        );
      
      default:
        return (
          <MessageBar messageBarType={MessageBarType.warning}>
            Unknown view type: {viewType}
          </MessageBar>
        );
    }
  }

  public render(): React.ReactElement<IAccordionTabsProps> {
    const { sections, isLoading, error } = this.state;

    return (
      <div className={styles.accordionTabsComponent}>
        {error && this.renderErrorState()}
        
        {isLoading && this.renderLoadingState()}
        
        {!isLoading && !error && (
          <div>
            {sections.length === 0 && this.props.displayMode !== 2 /* Edit Mode */ ? (
              this.renderEmptyState()
            ) : (
              <div className={styles.contentContainer}>
                {this.renderContent()}
              </div>
            )}
          </div>
        )}
      </div>
    );
  }
}
