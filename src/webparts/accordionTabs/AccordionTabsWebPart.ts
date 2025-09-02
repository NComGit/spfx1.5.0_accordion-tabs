import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  IPropertyPaneField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as React from 'react';
import * as ReactDom from 'react-dom';

import { IAccordionTabsWebPartProps, ViewType, ISection } from './models/IAccordionTabsModels';
import { AccordionTabsComponent } from './components/AccordionTabsComponent';
import { IAccordionTabsProps } from './models/IAccordionTabsModels';

import * as strings from 'AccordionTabsWebPartStrings';
import styles from './AccordionTabsWebPart.module.scss';

export default class AccordionTabsWebPart extends BaseClientSideWebPart<IAccordionTabsWebPartProps> {

  private _isDirty: boolean = false;

  public render(): void {
    const element: React.ReactElement<IAccordionTabsProps> = React.createElement(
      AccordionTabsComponent,
      {
        viewType: this.properties.viewType || ViewType.Accordion,
        sections: this.properties.sections || [],
        displayMode: this.displayMode,
        onSectionsChanged: this.onSectionsChanged.bind(this),
        onConfigureClick: this.onConfigureClick.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private onSectionsChanged(sections: ISection[]): void {
    // Update web part properties
    this.properties.sections = sections;
    this._isDirty = true;
    
    // Re-render the web part
    this.render();
    
    // Mark as dirty for SharePoint to save changes
    this.context.propertyPane.refresh();
  }

  private onConfigureClick(): void {
    // Open the property pane
    this.context.propertyPane.open();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getDataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const viewTypeOptions = [
      { key: ViewType.Accordion, text: strings.ViewTypeAccordion },
      { key: ViewType.Tabs, text: strings.ViewTypeTabs }
    ];

    // Build sections management fields
    const sectionFields: IPropertyPaneField<any>[] = [];
    
    // View type selector
    sectionFields.push(
      PropertyPaneChoiceGroup('viewType', {
        label: strings.ViewTypeLabel,
        options: viewTypeOptions
      })
    );

    // Add section button
    sectionFields.push(
      PropertyPaneButton('addSection', {
        text: strings.AddSectionButton,
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.onAddSection.bind(this)
      })
    );

    // Sections management
    if (this.properties.sections && this.properties.sections.length > 0) {
      const sortedSections = [...this.properties.sections].sort((a, b) => a.order - b.order);
      
      sortedSections.forEach((section, index) => {
        sectionFields.push(
          PropertyPaneTextField(`section_${section.id}_title`, {
            label: `${strings.SectionTitleLabel} ${index + 1}`,
            value: section.title,
            onGetErrorMessage: this.validateSectionTitle.bind(this),
            deferredValidationTime: 500
          })
        );

        // Add management buttons for each section
        sectionFields.push(
          PropertyPaneButton(`editSection_${section.id}`, {
            text: `${strings.EditSectionButton} ${index + 1}`,
            buttonType: PropertyPaneButtonType.Normal,
            onClick: () => this.onEditSection(section.id)
          })
        );

        sectionFields.push(
          PropertyPaneButton(`deleteSection_${section.id}`, {
            text: `${strings.DeleteSectionButton} ${index + 1}`,
            buttonType: PropertyPaneButtonType.Normal,
            onClick: () => this.onDeleteSection(section.id)
          })
        );

        if (index > 0) {
          sectionFields.push(
            PropertyPaneButton(`moveUpSection_${section.id}`, {
              text: `${strings.MoveUpButton} ${index + 1}`,
              buttonType: PropertyPaneButtonType.Normal,
              onClick: () => this.onMoveSectionUp(section.id)
            })
          );
        }

        if (index < sortedSections.length - 1) {
          sectionFields.push(
            PropertyPaneButton(`moveDownSection_${section.id}`, {
              text: `${strings.MoveDownButton} ${index + 1}`,
              buttonType: PropertyPaneButtonType.Normal,
              onClick: () => this.onMoveSectionDown(section.id)
            })
          );
        }
      });
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.ViewConfigurationGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('viewType', {
                  label: strings.ViewTypeLabel,
                  options: viewTypeOptions
                })
              ]
            },
            {
              groupName: strings.SectionsGroupName,
              groupFields: sectionFields
            }
          ]
        }
      ]
    };
  }

  private validateSectionTitle(value: string): string {
    if (!value || value.trim().length === 0) {
      return strings.SectionTitleRequiredError;
    }
    
    if (value.trim().length > 100) {
      return strings.SectionTitleLengthError;
    }
    
    return '';
  }

  private onAddSection(): void {
    const newSection: ISection = {
      id: this.generateSectionId(),
      title: strings.DefaultSectionTitle,
      content: '',
      order: this.properties.sections ? this.properties.sections.length : 0
    };

    if (!this.properties.sections) {
      this.properties.sections = [];
    }

    this.properties.sections.push(newSection);
    this._isDirty = true;
    
    // Refresh property pane and re-render
    this.context.propertyPane.refresh();
    this.render();
  }

  private onEditSection(sectionId: string): void {
    // This will be handled by the React components
    // The property pane is mainly for basic management
    console.log('Edit section:', sectionId);
  }

  private onDeleteSection(sectionId: string): void {
    if (!this.properties.sections) return;

    this.properties.sections = this.properties.sections.filter(s => s.id !== sectionId);
    this._isDirty = true;
    
    // Refresh property pane and re-render
    this.context.propertyPane.refresh();
    this.render();
  }

  private onMoveSectionUp(sectionId: string): void {
    if (!this.properties.sections) return;

    const sortedSections = [...this.properties.sections].sort((a, b) => a.order - b.order);
    
    // Find current index
    let currentIndex = -1;
    for (let i = 0; i < sortedSections.length; i++) {
      if (sortedSections[i].id === sectionId) {
        currentIndex = i;
        break;
      }
    }
    
    if (currentIndex <= 0) return;

    // Swap orders
    const temp = sortedSections[currentIndex].order;
    sortedSections[currentIndex].order = sortedSections[currentIndex - 1].order;
    sortedSections[currentIndex - 1].order = temp;

    this._isDirty = true;
    this.context.propertyPane.refresh();
    this.render();
  }

  private onMoveSectionDown(sectionId: string): void {
    if (!this.properties.sections) return;

    const sortedSections = [...this.properties.sections].sort((a, b) => a.order - b.order);
    
    // Find current index
    let currentIndex = -1;
    for (let i = 0; i < sortedSections.length; i++) {
      if (sortedSections[i].id === sectionId) {
        currentIndex = i;
        break;
      }
    }
    
    if (currentIndex === -1 || currentIndex >= sortedSections.length - 1) return;

    // Swap orders
    const temp = sortedSections[currentIndex].order;
    sortedSections[currentIndex].order = sortedSections[currentIndex + 1].order;
    sortedSections[currentIndex + 1].order = temp;

    this._isDirty = true;
    this.context.propertyPane.refresh();
    this.render();
  }

  private generateSectionId(): string {
    return 'section-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    // Handle section title changes - using indexOf for TypeScript 2.4.2 compatibility
    if (propertyPath.indexOf('section_') === 0 && propertyPath.indexOf('_title') === propertyPath.length - 6) {
      const sectionIdMatch = propertyPath.match(/section_(.+)_title/);
      if (sectionIdMatch && this.properties.sections) {
        const sectionId = sectionIdMatch[1];
        // Use manual loop for TypeScript 2.4.2 compatibility instead of find()
        let section: ISection = null;
        for (let i = 0; i < this.properties.sections.length; i++) {
          if (this.properties.sections[i].id === sectionId) {
            section = this.properties.sections[i];
            break;
          }
        }
        if (section) {
          section.title = newValue;
          this._isDirty = true;
        }
      }
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected getDisableReactivePropertyChanges(): boolean {
    return false;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this._isDirty = false;
  }

  protected getCanOpenPropertyPane(): boolean {
    return true;
  }
}
