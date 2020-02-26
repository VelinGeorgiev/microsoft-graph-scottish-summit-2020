import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'AssignAManagerWebPartStrings';
import { AadHttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';

export interface IAssignAManagerWebPartProps {
  description: string;
}

export default class AssignAManagerWebPart extends BaseClientSideWebPart <IAssignAManagerWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <input type="input" id="employee" placeholder="Email of the employee" autocomplete="on" name="email" />
      <input type="input" id="manager" placeholder="Email of the manager" autocomplete="on" name="email" />
      <button id="assign">Assign</button>
      <pre id="result"></pre>`;

    const assignButton = document.getElementById('assign');
    assignButton.addEventListener('click', () => {

      const manager = document.getElementById('manager') as HTMLInputElement;
      const employee = document.getElementById('employee') as HTMLInputElement;

      this.context.aadHttpClientFactory
        .getClient('https://graph.microsoft.com')
        .then((client: AadHttpClient) => {

          const requestHeaders: Headers = new Headers();
          requestHeaders.append('Content-type', 'application/json');
          requestHeaders.append('Cache-Control', 'no-cache');

          const options: IHttpClientOptions = {
            headers: requestHeaders,
            body: `{ "@odata.id": "https://graph.microsoft.com/v1.0/users/${encodeURIComponent(employee.value)}" }`
          };

          client.post(`https://graph.microsoft.com/users/${encodeURIComponent(manager.value)}/manager/$ref`, AadHttpClient.configurations.v1, options)
          .then((response: HttpClientResponse) => {

            return response.json()
          })
          .then((responseJSON: JSON) => {

            const result = document.getElementById('result');
            result.innerHTML = JSON.stringify(responseJSON, null, 2);

          })
          .catch((error: any) => {

            const result = document.getElementById('result');
            result.innerHTML = JSON.stringify(error, null, 2);
            
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
