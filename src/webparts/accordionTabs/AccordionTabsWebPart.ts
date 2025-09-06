import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as React from 'react';
import * as ReactDom from 'react-dom';

import { IAccordionTabsWebPartProps, ViewType, ISection, AccordionDefaultExpanded, TabsDefaultActive } from './models/IAccordionTabsModels';
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
        onConfigureClick: this.onConfigureClick.bind(this),
        // Accordion settings
        accordionDefaultExpanded: this.properties.accordionDefaultExpanded || AccordionDefaultExpanded.First,
        accordionChosenSection: this.properties.accordionChosenSection || "",
        // Tabs settings
        tabsDefaultActive: this.properties.tabsDefaultActive || TabsDefaultActive.First,
        tabsChosenTab: this.properties.tabsChosenTab || ""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Ensure default values are set
    if (!this.properties.viewType) {
      this.properties.viewType = ViewType.Accordion;
    }
    if (!this.properties.accordionDefaultExpanded) {
      this.properties.accordionDefaultExpanded = AccordionDefaultExpanded.First;
    }
    if (this.properties.accordionChosenSection === undefined) {
      this.properties.accordionChosenSection = "";
    }
    if (this.properties.tabsDefaultActive === undefined) {
      this.properties.tabsDefaultActive = TabsDefaultActive.First;
    }
    if (this.properties.tabsChosenTab === undefined) {
      this.properties.tabsChosenTab = "";
    }

    const viewTypeOptions = [
      { key: ViewType.Accordion, text: strings.ViewTypeAccordion },
      { key: ViewType.Tabs, text: strings.ViewTypeTabs }
    ];

    // Build conditional fields based on view type
    const conditionalFields = [];

    if (this.properties.viewType === ViewType.Accordion) {
      // Accordion-specific settings
      const accordionExpandedOptions = [
        { key: AccordionDefaultExpanded.None, text: "No layers" },
        { key: AccordionDefaultExpanded.First, text: "First layer" },
        { key: AccordionDefaultExpanded.All, text: "All layers" },
        { key: AccordionDefaultExpanded.Chosen, text: "Chosen layer" }
      ];

      conditionalFields.push(
        PropertyPaneDropdown('accordionDefaultExpanded', {
          label: "Default expanded layers",
          options: accordionExpandedOptions,
          selectedKey: this.properties.accordionDefaultExpanded || AccordionDefaultExpanded.First
        })
      );

      // Show section chooser only when "Chosen" is selected
      if (this.properties.accordionDefaultExpanded === AccordionDefaultExpanded.Chosen && this.properties.sections && this.properties.sections.length > 0) {
        const sectionOptions = this.properties.sections
          .sort((a, b) => a.order - b.order)
          .map((section, index) => ({
            key: section.id,
            text: section.title || `Section ${index + 1}`
          }));

        conditionalFields.push(
          PropertyPaneDropdown('accordionChosenSection', {
            label: "Choose section to expand",
            options: sectionOptions,
            selectedKey: this.properties.accordionChosenSection || ""
          })
        );
      }
    } else if (this.properties.viewType === ViewType.Tabs) {
      // Tabs-specific settings
      const tabsDefaultActiveOptions = [
        { key: TabsDefaultActive.First, text: "First Tab" },
        { key: TabsDefaultActive.Last, text: "Last Tab" },
        { key: TabsDefaultActive.Chosen, text: "Chosen Tab" }
      ];

      conditionalFields.push(
        PropertyPaneDropdown('tabsDefaultActive', {
          label: "Default active tab",
          options: tabsDefaultActiveOptions,
          selectedKey: this.properties.tabsDefaultActive || TabsDefaultActive.First
        })
      );

      // Show tab chooser only when "Chosen" is selected
      if (this.properties.tabsDefaultActive === TabsDefaultActive.Chosen && this.properties.sections && this.properties.sections.length > 0) {
        const tabOptions = this.properties.sections
          .sort((a, b) => a.order - b.order)
          .map((section, index) => ({
            key: section.id,
            text: section.title || `Tab ${index + 1}`
          }));

        conditionalFields.push(
          PropertyPaneDropdown('tabsChosenTab', {
            label: "Choose tab to display",
            options: tabOptions,
            selectedKey: this.properties.tabsChosenTab || ""
          })
        );
      }
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
                }),
                ...conditionalFields
              ]
            }
          ]
        }
      ]
    };
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
