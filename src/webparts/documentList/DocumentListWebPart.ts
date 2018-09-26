import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
} from '@microsoft/sp-webpart-base';
/** Property Pane Controls Reference */
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout, PropertyFieldNumber } from '@pnp/spfx-property-controls';
import { PropertyFieldChoiceGroupWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldChoiceGroupWithCallout';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldDropdownWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldDropdownWithCallout';
/** SP PnP Reference */
import { sp } from '@pnp/sp';


import * as strings from 'DocumentListWebPartStrings';
import DocumentList from './components/documentList/DocumentList';
import { IDocumentListProps } from './components/documentList/IDocumentListProps';

export interface IDocumentListWebPartProps {
  title: string;
  docLibURL: string;
  layoutType: string;
  dateFormat: string;
  showFolder: boolean;
  itemsPerPage: number;
}

export default class DocumentListWebPart extends BaseClientSideWebPart<IDocumentListWebPartProps> {

  protected onInit(): Promise<void> {
    // Setup the PnP Context
    sp.setup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IDocumentListProps> = React.createElement(
      DocumentList,
      {
        currentContext: this.context,
        displayMode: this.displayMode,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        serviceScope: this.context.serviceScope,
        title: this.properties.title,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        doclibUrl: this.properties.docLibURL,
        layoutType: this.properties.layoutType,
        dateFormat: this.properties.dateFormat,
        showFolder: this.properties.showFolder,
        itemsPerPage: this.properties.itemsPerPage
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldTextWithCallout('docLibURL', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'docLibURLId',
                  label: strings.DocLibraryFieldLabel,
                  calloutContent: React.createElement('div', {}, strings.DocLibraryFieldCalloutContent),
                  value: this.properties.docLibURL
                }),
                PropertyFieldChoiceGroupWithCallout('layoutType', {
                  calloutContent: React.createElement('div', {}, strings.LayoutTypeFieldCalloutContent),
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'layoutTypeFieldId',
                  label: strings.LayoutTypeFieldLabel,
                  options: [
                    {
                      key: 'box',
                      text: 'Box',
                      checked: this.properties.layoutType === 'box'
                    }, {
                      key: 'list',
                      text: 'List',
                      checked: this.properties.layoutType === 'list'
                    },
                    {
                      key: 'dccl',
                      text: 'Document Card (Compact Layout)',
                      checked: this.properties.layoutType === 'dccl'
                    }
                  ],
                }),
                PropertyFieldChoiceGroupWithCallout('dateFormat', {
                  calloutContent: React.createElement('div', {}, strings.DateFormatFieldCalloutContent),
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'dateFormatField',
                  label: strings.DateFormatFieldLabel,
                  options: [
                    {
                      key: 'DD/MM/YYYY',
                      text: 'DD/MM/YYYY (31/01/2018)',
                      checked: this.properties.dateFormat === 'DD/MM/YYYY',
                    },
                    {
                      key: 'MM/DD/YYYY',
                      text: 'MM/DD/YYYY (01/31/2018)',
                      checked: this.properties.dateFormat === 'MM/DD/YYYY',
                    },
                    {
                      key: 'DD MMM YYYY',
                      text: 'DD MMM YYYY (31 Jan 2018)',
                      checked: this.properties.dateFormat === 'DD MMM YYYY',
                    },
                    {
                      key: 'MMM DD YYYY',
                      text: 'MMM DD YYYY (Jan 31 2018)',
                      checked: this.properties.dateFormat === 'MMM DD YYYY',
                    }
                  ]
                }),
                PropertyFieldToggleWithCallout('showFolder', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'showFolderFieldId',
                  label: strings.ShowFoldersFieldLabel,
                  calloutContent: React.createElement('div', {}, strings.ShowFoldersFieldCalloutContent),
                  onText: 'ON',
                  offText: 'OFF',
                  checked: this.properties.showFolder
                }),
                PropertyFieldNumber('itemsPerPage', {
                  label: strings.ItemsPerPageLabel,
                  description: strings.ItemsPerPageDescription,
                  key: "itemsPerPageField",
                  minValue: 1,
                  maxValue: 12,
                  value: this.properties.itemsPerPage
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
