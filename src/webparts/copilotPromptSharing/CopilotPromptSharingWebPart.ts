import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CopilotPromptSharingWebPartStrings';
import CopilotPromptSharing from './components/CopilotPromptSharing';
import { ICopilotPromptSharingProps } from './components/ICopilotPromptSharingProps';

export interface ICopilotPromptSharingWebPartProps {
  description: string;
}

export default class CopilotPromptSharingWebPart extends BaseClientSideWebPart<ICopilotPromptSharingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICopilotPromptSharingProps> =
      React.createElement(
        CopilotPromptSharing,
        {
          description: this.properties.description,
          context: this.context
        }
      );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
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