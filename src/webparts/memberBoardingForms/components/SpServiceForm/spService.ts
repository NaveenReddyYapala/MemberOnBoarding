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
    "LegalTeam": "AssignedToLegalTeam",
    "CheckerTeam": "AssignedToCheckerTeam",
    "FinanceTeam": "AssignedToFinanceTeam",
    "IdsupportTeam": "AssignedToIdsupportTeam",
    "FTPSupportTeam": "AssignedToFTPSupportTeam",
    "MOBClouserTeam": "AssignedToMOBClouserTeam"
  };

  public async getUserGroups(): Promise<string[]> {
    const groups = await sp.web.currentUser.groups();
    return groups.map(g => g.Title);
  }

  public async getItemById(listName: string, id: any): Promise<any> {
    const item = await sp.web.lists.getByTitle(listName).items.getById(id).select("Id", "Title", "Status").get();
    return item;
  }

//   public async getItemById(): Promise<any> {
//   // Get query string
//   const params = new URLSearchParams(window.location.search);
//   const idParam = params.get("ID");

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
//   return item;
// }

  

  public getGroupStatusMap() {
    return this.groupStatusMap;
  }
}
