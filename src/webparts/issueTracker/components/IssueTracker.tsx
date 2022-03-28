import * as React from 'react';
import styles from './IssueTracker.module.scss';
import { IIssueTrackerProps } from './IIssueTrackerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from "./../pnpjsConfig";
import { SPFI, spfi, } from "@pnp/sp";
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';


export default class IssueTracker extends React.Component<IIssueTrackerProps, {
  items: any[]
}> {
  private _sp: SPFI;

  constructor(props) {
    super(props);
    this._sp = getSP();

    this.state = {
      items: undefined
    }
  }

  public componentDidMount(): void {
    this.getListItems();
  }

  private getListItems = async () => {
    try {
      const items: any[] = await this._sp.web.lists.getByTitle('Issue tracker list').items();
      console.log("items - ", items);
      this.setState({
        items: items
      });
    }
    catch (error) {
      console.error('Error getting items - ', error);
    }
  }

  private listColumns: IColumn[] = [
    { key: 'title', name: 'Name', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'description', name: 'Description', fieldName: 'Description', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 50, maxWidth: 100, isResizable: true },
    { key: 'priority', name: 'Priority', fieldName: 'Priority', minWidth: 50, maxWidth: 100, isResizable: true },
  ];

  public render(): React.ReactElement<IIssueTrackerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.issueTracker} ${hasTeamsContext ? styles.teams : ''}`}>
        {this.state.items && 
          <DetailsList items={this.state.items} columns={this.listColumns} />
        }
      </section>
    );
  }
}
