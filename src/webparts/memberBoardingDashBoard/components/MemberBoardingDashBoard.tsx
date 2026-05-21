import * as React from 'react';
import styles from './MemberBoardingDashBoard.module.scss';
import { IMemberBoardingDashBoardProps } from './IMemberBoardingDashBoardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices from './SpServiceDashboard/spService';
import { IconButton } from "office-ui-fabric-react/lib/Button";

export interface IBoardingItem {
  Id: number;
  Title: string;
  CreditInstitutionName: string;
  // KOMGrouping: {
  //   Id: number;
  //   Title: string;
  // };
  NewMemberCode: string;
  MemberShortCode: string;
  Status: string;
}

export interface IMemberBoardingDashBoardState {
  getItems: IBoardingItem[];
}

export default class MemberBoardingDashBoard extends React.Component<IMemberBoardingDashBoardProps, IMemberBoardingDashBoardState> {
  private spService: spservices;
  constructor(props: IMemberBoardingDashBoardProps) {
    super(props);
    this.state = {
      getItems: []
    };
    this.spService = new spservices(this.props.context);
  }
  public async componentDidMount() {
    try {
      const style = document.createElement('style');
      style.innerHTML = `    
           .banner_dbeb3a78 { display: none; }
      .pagetitle_f969f8ac { display: none; }
            `;
      document.head.appendChild(style);
      //this._loadDropdownData();
      const getItems = await this.spService.getFilteredItems();
      this.setState({ getItems });
    } catch (error) {
      console.log("error", error)
    }
  }
  private onEditClick(item: IBoardingItem): void {
    if (item.Status === "Draft" || item.Status === "Assigned to Maker") {
      window.location.href = `/sites/NaveenReddy/SitePages/MemberOnBoardingNewForm.aspx?ItemId=${item.Id}`;
    } else {
      window.location.href = `/sites/NaveenReddy/SitePages/MemberBoardingForm.aspx?ItemId=${item.Id}`;
    }
  }

  public render(): React.ReactElement<IMemberBoardingDashBoardProps> {
    return (
      <div className={styles.memberBoardingDashBoard} >
        <table>
          <thead>
            <tr>
              <th>Edit</th>
              <th>ID</th>
              <th>Title</th>
              <th>CreditInstitutionName</th>
              <th>NewMemberCode</th>
              <th>MemberShortCode</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {this.state.getItems && this.state.getItems.map(item => (
              <tr key={item.Id}>
                <td>
                  <img alt="edit" src="/_layouts/15/images/edititem.gif?rev=43" 
                  onClick={() => this.onEditClick(item)} /></td>
{/*                                         
                <td>
                  <IconButton
                    iconProps={{ iconName: "Edit" }}
                    title="Edit"
                    ariaLabel="Edit"
                    onClick={() => this.onEditClick(item)}
                  />
                </td> */}
                <td>{item.Id}</td>
                <td>{item.Title}</td>
                <td>{item.CreditInstitutionName}</td>
                <td>{item.NewMemberCode}</td>
                <td>{item.MemberShortCode}</td>
                <td>{item.Status}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div >
    );
  }
}
