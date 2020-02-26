import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SearchForAManagerWebPartStrings';
import { MSGraphClient } from '@microsoft/sp-http';


export interface ISearchForAManagerWebPartProps {
  description: string;
}

export default class SearchForAManagerWebPart extends BaseClientSideWebPart <ISearchForAManagerWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div>
      <input type="text" id="email" name="email" autocomplete="on" />
      <button id="search">Search</button>
      <div>Direct Manager:</div>
      <pre id="result"></pre> 
    </div>`;

    const searchButton = document.getElementById('search');
    searchButton.addEventListener('click', () => {

      const email = document.getElementById('email') as HTMLInputElement;

      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
            .api(`/users/${encodeURIComponent(email.value)}/manager/`)
            .version('v1.0')
            .select('displayName,mail,userPrincipalName')
            .get((err, res) => {

              const result = document.getElementById('result');
              result.innerHTML = JSON.stringify((err || res), null, 2);
              
            });
        });
    });
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
