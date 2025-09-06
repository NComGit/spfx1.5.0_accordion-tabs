import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup
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
            }
          ]
        }
      ]
    };
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
