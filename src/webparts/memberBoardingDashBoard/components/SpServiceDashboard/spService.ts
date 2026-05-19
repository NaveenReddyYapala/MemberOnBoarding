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

    const items = await sp.web.lists.getByTitle("Membership OnBoarding Request").items
      .filter(filterQuery)
      .select("Id", "Title", "AddressSameAsInProfile", "AssetSize", "BillingAddress1", "BillingAddress2", "BillingAddressSameAsInProfile", "BillingCity", "BillingContactName", "BillingEmailAddress", "BillingPinCode", "BillingSPOC", "BillingTelephoneNumber", "BusinessAddress1", "BusinessAddress2", "BusinessAddressPinCode", "BusinessAddressSameAsInProfile", "BusinessCIty", "BusinessContactName", "CertifiedLicenseIssuedByRBISigne", "CertifiedROCCertificate", "CreditInstitutionName", "CreditRiskContact", "DataAddress1", "DataAddress2", "DataAddressSameAsInProfile", "DataCity", "DataContactName", "DataEmailAddress", "DataPinCode", "DataTelephoneNumber", "DateOfMembership", "Devation_x002f_ExpectionApproval", "DuplicateName_x002f_Code_x002f_S", "EmailAddress", "GSTBillingState/Id", "GSTBillingState/Title", "GSTNumberExists", "ITSPOC", "KAMId/Id", "KAMId/Title", "KAMId/EMail", "KOMGrouping/Id", "KOMGrouping/Title", "LandlineNumberOffice", "LatestBalancesheet_x002f_AnnualR", "LetterApplicationProspectsLetter", "LetterAuthorityIssued", "LicenseCancellation_x002f_OtherT", "MCAWebsiteChecked", "MembershipApplicationForm", "MemberShortCode", "Mobile", "NameOfCoreBanking", "NewBillingState/Id", "NewBillingState/Title", "NewBusinessAddressState/Id", "NewBusinessAddressState/Title", "NewDataState/Id", "NewDataState/Title", "NewMemberCode", "NewNominatedNodalState/Id", "NewNominatedNodalState/Title", "NewRegisteredNumber", "NewRegisteredOfficeState/Id", "NewRegisteredOfficeState/Title", "NominatedCity", "NominatedNodalAddress1", "NominatedNodalAddress2", "NominatedNodalEmailAddress", "NominatedNodalFaxNumber", "NominatedNodalLandlineNumber", "NominatedNodalOfficerName", "NominatedNodalPincode", "OperatingRuleBook", "PAN", "PaymentAnnualFee", "PaymentOfMembershipFee", "PricingAnnexure", "Rate_x0020_Type", "RBI_x002f_RegulatoryListChecked", "RegisteredCity", "RegisteredOfficeAddress1", "RegisteredOfficeAddress2", "RegisteredOfficePinCode", "Remarks", "Senddocumentsby", "SiteVerificationreport", "TAN", "TypeofPricing",  "TVRVerification", "TVRComment", "WebSearchPerformed", "MCACheck", "RBICheck", "OFACCheck", "LitigationComment", "Status", "Author/Id", "Author/Title", "Author/EMail")
      .expand("Author", "GSTBillingState", "KAM", "KOMGrouping", "NewBillingState", "NewBusinessAddressState", "NewDataState", "NewNominatedNodalState", "NewRegisteredOfficeState")
      .get();

    return items;

  }

}
