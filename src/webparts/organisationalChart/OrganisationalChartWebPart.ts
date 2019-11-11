import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'OrganisationalChartWebPartStrings';
import OrganisationalChart from './components/OrganisationalChart';
import { IOrganisationalChartProps } from './components/IOrganisationalChartProps';
import TreeOrgChart from './components/OrganisationalChart';
// import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { setup as pnpSetup } from '@pnp/common';
import { PageContext } from '@microsoft/sp-page-context';

export interface IOrganisationalChartWebPartProps {
  description: string;
  title: string;
  listName:string;
  currentUserTeam: boolean;
  maxLevels: number;
  numberOfItems: number;
  pageContext: PageContext;// here we passing page context
}

export default class OrganisationalChartWebPart extends BaseClientSideWebPart<IOrganisationalChartWebPartProps> {
  public onInit(): Promise<void> {

    pnpSetup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }
  public render(): void {
    const element: React.ReactElement<IOrganisationalChartProps > = React.createElement(
      OrganisationalChart,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        numberOfItems: this.properties.numberOfItems,
        pageContext: this.context.pageContext
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
               
                // PropertyPaneTextField('listName', {
                  
                //   label: strings.ListNameFieldLabel,
                  
                // }),
                PropertyPaneSlider('numberOfItems', {
                  
                  label: strings.numberOfItemsFieldLabel,
                  min: 1,
                  max: 6,
                  step: 1
                }),

                // PropertyFieldNumber("maxLevels", {
                //   key: "numberValue",
                //   label: strings.MaxLevels,
                //   description: strings.MaxLevels,
                //   value: this.properties.maxLevels,
                //   maxValue: 10,
                //   minValue: 1,
                //   disabled: false
                // })
              ]
            }
          ]
        }
      ]
    };
  }
}
