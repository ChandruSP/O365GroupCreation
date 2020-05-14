import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CreateO365GroupsWebPartStrings';
import CreateO365Groups from './components/CreateO365Groups';
import { ICreateO365GroupsProps } from './components/ICreateO365GroupsProps';

import { MSGraphClient, HttpClient } from '@microsoft/sp-http';

export interface ICreateO365GroupsWebPartProps {
  description: string;
}

export default class CreateO365GroupsWebPart extends BaseClientSideWebPart<ICreateO365GroupsWebPartProps> {

  public render(): void {
    var email = this.context.pageContext.user.email;
    var _httpClient = this.context.httpClient;
    var currentContext = this.context;

    this.context.msGraphClientFactory.getClient()
      .then((_graphClient: MSGraphClient): void => {
        const element: React.ReactElement<ICreateO365GroupsProps> = React.createElement(
          CreateO365Groups,
          {
            graphClient: _graphClient,
            userEmail: email,
            httpClient: _httpClient,
            context: currentContext
          }
        );
        ReactDom.render(element, this.domElement);
      });

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
