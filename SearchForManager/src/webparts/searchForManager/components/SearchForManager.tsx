import * as React from 'react';
import { ISearchForManagerProps } from './ISearchForManagerProps';
import { MSGraphClient } from '@microsoft/sp-http';

interface ISearchForManagerState {
  result: any;
}

export default class SearchForManager extends React.Component<ISearchForManagerProps, ISearchForManagerState> {

  constructor(props: ISearchForManagerProps) {
    super(props);
    this.state = { result: '' };
  }

  public render(): React.ReactElement<ISearchForManagerProps> {
    return (
      <div>
        <input type="search" id="employee" />
        <button onClick={this.search.bind(this)}>Search</button>
        <pre>{JSON.stringify(this.state.result, null, 2)}</pre>
      </div>
    );
  }

  private search(): void {
    const employee: HTMLInputElement = document.getElementById("employee") as HTMLInputElement;
    if (employee.value) {

      this.props.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
            .api(`/users/${employee.value}/manager/`)
            .version("v1.0")
            .select("displayName,mail,userPrincipalName")
            .get((err, res) => {

              this.setState({ result: err || res.displayName });
            });
        });
    }
  }
}