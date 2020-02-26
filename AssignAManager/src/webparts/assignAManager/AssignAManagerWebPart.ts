import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'AssignAManagerWebPartStrings';
import { MSGraphClient } from '@microsoft/sp-http';

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

      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {

          client
          .api(`/users/${encodeURIComponent(manager.value)}/manager/$ref`)
          .header('Content-Type', 'application/json')
          .put({ 
                "@odata.id": `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(employee.value)}` 
               }, 
              (err, res) => {

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
