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
import { IWebPartContext } from '@microsoft/sp-webpart-base'
import { IDigestCache, DigestCache } from '@microsoft/sp-http';

export interface IAdminFormWebPartProps {
  description: string;
  context: IWebPartContext
  queryString: string;
}

export default class AdminFormWebPart extends BaseClientSideWebPart<IAdminFormWebPartProps> {
  public digest: string = "";

  public constructor(context: IWebPartContext) {
    super();
  }

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);
      digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl).then((digest: string): void => {
        // use the digest here
        this.digest = digest;
        resolve();
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IAdminFormProps> = React.createElement(
      AdminForm,
      {
        digest: this.digest,
        context: this.context,
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        queryString: this.properties.queryString,
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
