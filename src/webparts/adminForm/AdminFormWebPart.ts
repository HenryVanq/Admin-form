import * as React from 'react';
import * as ReactDom from 'react-dom';
import { default as pnp, ItemUpdateResult, Web, Item, sp } from "sp-pnp-js";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'AdminFormWebPartStrings';
import AdminForm from './components/AdminForm';
import { IAdminFormProps } from './components/IAdminFormProps';

export interface IAdminFormWebPartProps {
  description: string;
}

export default class AdminFormWebPart extends BaseClientSideWebPart<IAdminFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAdminFormProps> = React.createElement(
      AdminForm,
      {
        context: this.context,
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl,
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
