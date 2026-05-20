import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web, SiteGroups, Item } from '@pnp/sp/presets/all';
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export default class spservices {
  private context: any;

  constructor(context: any) {
    this.context = context;
    sp.setup({ sp: { baseUrl: this.context.pageContext.web.absoluteUrl } });
  }

  private groupStatusMap: { [key: string]: string } = {
    "LegalTeam": "Assigned To Legal",
    "CheckerTeam": "Assigned To Checker",
    "Member On-Boarding Finance Team": "Assigned To Finance",
    "ID Support Team": "Assigned To IDSupport",
    "FTP Support Team": "Assigned To FTPSupport",
    "Membership On-Boarding Closure": "Assigned To Closure"
  };

  public async getUserGroups(): Promise<string[]> {
    const groups = await sp.web.currentUser.groups();
    return groups.map(g => g.Title);
  }

  public async getAllowedStatuses(): Promise<string[]> {
    const userGroups = await this.getUserGroups();
    const allowedStatuses: string[] = [];

    userGroups.forEach(group => {
      if (this.groupStatusMap[group]) {
        allowedStatuses.push(this.groupStatusMap[group]);
      }
    });

    return allowedStatuses;
  }

  // public async getFilteredItems(): Promise<any[]> {
  //   const allowedStatuses = await this.getAllowedStatuses();

  //   if (allowedStatuses.length === 0) {
  //     return [];
  //   }

  //   const filterQuery = allowedStatuses.map(s => `Status eq '${s}'`).join(" or ");
  //   const items = await sp.web.lists.getByTitle("Membership OnBoarding Request").items
  //     .filter(filterQuery)
  //     .select("Id", "Title", "Status")
  //     .get();

  //   return items;
  // }
  public async getFilteredItems(): Promise<any[]> {
    const allowedStatuses = await this.getAllowedStatuses();
    const currentUser = await sp.web.currentUser.get();

    // Build filter for group-based statuses
    const statusFilter = allowedStatuses.length > 0
      ? allowedStatuses.map(s => `Status eq '${s}'`).join(" or ")
      : "";

    // Build filter for items created by current user with Status = Draft
    const draftFilter = `(Author/Id eq ${currentUser.Id} and Status eq 'Draft')`;

    // Combine filters
    let filterQuery = "";
    if (statusFilter) {
      filterQuery = `${statusFilter} or ${draftFilter}`;
    } else {
      filterQuery = draftFilter;
    }

    const items = await sp.web.lists.getByTitle("Membership On-Boarding Request").items
      .filter(filterQuery)
      .select("Id", "Title", "CreditInstitutionName", "MemberShortCode",  "NewMemberCode", "Status")     
      .get();

    return items;

  }

}
