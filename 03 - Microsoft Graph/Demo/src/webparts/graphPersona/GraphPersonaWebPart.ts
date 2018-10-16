import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphPersonaWebPartStrings';
import GraphPersona from './components/GraphPersona';
import { IGraphPersonaService } from './components/IGraphPersonaService';
import { GraphPersonaService } from './components/GraphPersonaService';
import { IGraphPersonaProps } from './components/IGraphPersonaProps';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphPersonaWebPartProps {
  description: string;
}

export default class GraphPersonaWebPart extends BaseClientSideWebPart<IGraphPersonaWebPartProps> {

  public async render(): Promise<void> {
    var service : IGraphPersonaService;
    if (Environment.type == EnvironmentType.Local) {
        service = new (await import(/* webpackChunkName: 'dummyservice' */'./components/GraphPersonaService.dummy')).GraphPersonaService();
    } else {
        var client = await this.context.msGraphClientFactory.getClient();
        service = new GraphPersonaService(client);
    }
    const element: React.ReactElement<IGraphPersonaProps> = React.createElement(
        GraphPersona,
        {
            service: service
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
