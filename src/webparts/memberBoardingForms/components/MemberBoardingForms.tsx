import * as React from 'react';
import styles from './MemberBoardingForms.module.scss';
import { IMemberBoardingFormsProps } from './IMemberBoardingFormsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";
import { sp } from "@pnp/sp";
import spservices from './SpServiceForm/spService';
import LegalForm from './Forms/LegalForm';
import MOClouserForm from './Forms/MOClouserForm';
import IDSupportForm from './Forms/IDSupportForm';
import FTPSupportForm from './Forms/FTPSupportForm';
import FinanceForm from './Forms/FinanceForm';
import CheckerForm from './Forms/CheckerForm';

export interface IItem {
  Id: number;
  Title: string;
  Status: string;
}

export interface IMemberBoardingFormsState {
  item: IItem | null;
  userGroups: string[];
}

export default class MemberBoardingForms extends React.Component<IMemberBoardingFormsProps, IMemberBoardingFormsState> {
  private spService: spservices;
  constructor(props: IMemberBoardingFormsProps) {
    super(props);
    this.state = {
      item: null,
      userGroups: []
    };
    this.spService = new spservices(this.props.context);
  }
  public async componentDidMount() {
    try {
      const queryParams = new UrlQueryParameterCollection(window.location.href);
      const itemId = queryParams.getValue("ItemId");

      if (!itemId) return;

      const item = await this.spService.getItemById("Membership OnBoarding Request", itemId);
      //  const params = new URLSearchParams(window.location.search);
      //   const idParam = params.get("ItemId");
      
      //   if (!idParam) {
      //     throw new Error("No ID parameter found in URL");
      //   }
      
      //   const id = parseInt(idParam, 10);
      // console.log(id);
      //   const item = await sp.web.lists
      //     .getByTitle("Membership OnBoarding Request")
      //     .items.getById(id)
      //     .select("Id", "Title", "Status")
      //     .get();
      // console.log(item);
      const userGroups = await this.spService.getUserGroups();

      this.setState({ item, userGroups });
    } catch (error) {
      console.log("error", error)
    }
  }

  private renderForm() {
    const { item, userGroups } = this.state;
    if (!item) return null;

    const status = item.Status;
    const map = this.spService.getGroupStatusMap();

    // Check each group → status mapping
    if (userGroups.indexOf("LegalTeam") > -1 && status === map["LegalTeam"]) {
      return <LegalForm item={item} />;
    }
    
    if (userGroups.indexOf("CheckerTeam") > -1 && status === map["CheckerTeam"]) {
      return <CheckerForm item={item} />;
    }
    if (userGroups.indexOf("FinanceTeam") > -1 && status === map["FinanceTeam"]) {
      return <FinanceForm item={item} />;
    }
    if (userGroups.indexOf("IdsupportTeam") > -1 && status === map["IdsupportTeam"]) {
      return <IDSupportForm item={item} />;
    }
    if (userGroups.indexOf("FTPSupportTeam") > -1 && status === map["FTPSupportTeam"]) {
      return <FTPSupportForm item={item} />;
    }
    if (userGroups.indexOf("MOBClouserTeam") > -1 && status === map["MOBClouserTeam"]) {
      return <MOClouserForm item={item} />;
    }

    return <div>No form available for your group/status.</div>;
  }

  public render(): React.ReactElement<IMemberBoardingFormsProps> {
    return (
      <div className={styles.memberBoardingForms} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {this.renderForm()}
            </div>
          </div>
        </div>
      </div >
    );
  }
}
