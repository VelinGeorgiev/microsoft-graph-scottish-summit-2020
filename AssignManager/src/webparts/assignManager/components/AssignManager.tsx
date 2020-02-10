import * as React from 'react';
import { IAssignManagerProps } from './IAssignManagerProps';
import { MSGraphClient } from '@microsoft/sp-http';

interface IAssignManagerState {
  result: any;
}

export default class AssignManager extends React.Component<IAssignManagerProps, IAssignManagerState> {

  constructor(props: IAssignManagerProps) {
    super(props);
    this.state = { result: '' };
  }

  public render(): React.ReactElement<IAssignManagerProps> {
    return (
      <div>
        <input type="input" id="employee" placeholder="employee" />
        <input type="input" id="manager" placeholder="manager" />
        <button onClick={this.assignManager.bind(this)}>Assign</button>
        <pre>{JSON.stringify(this.state.result, null, 2)}</pre>
      </div>
    );
  }

  private assignManager(): void {
    const manager: HTMLInputElement = document.getElementById("manager") as HTMLInputElement;
    const employee: HTMLInputElement = document.getElementById("employee") as HTMLInputElement;

    if (manager.value && employee.value) {

      this.props.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          
          client
          .api(`/users/${manager.value}/manager/$ref`)
          .header("Content-Type", "application/json")
          .put({ "@odata.id": `https://graph.microsoft.com/v1.0/users/${employee.value}` }, (err, res) => {
            
            return this.setState({ result: err || 'DONE' });
          });
        });
    }
  }
}
