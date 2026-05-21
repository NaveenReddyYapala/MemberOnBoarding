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
    "ID Support Team": "Assigned To UAM",
    "FTP Support Team": "Assigned To UAM",
    // "ID Support Team": "Assigned To IDSupport",
    // "FTP Support Team": "Assigned To FTPSupport",
    "Membership On-Boarding Closure": "Assigned To Closure"
  };

  public async getUserGroups(): Promise<string[]> {
    const groups = await sp.web.currentUser.groups();
    return groups.map(g => g.Title);
  }
 public async GetStateOptions(): Promise<IDropdownOption[]> {
    const items = await sp.web.lists.getByTitle("MOBStates").items.select("Id, Title").get();
    return items.map((item: { Id: number; Title: string }) => ({
      key: item.Id,
      text: item.Title
    }));
  }
  public async GetKomGroupOptions(): Promise<IDropdownOption[]> {
    const items = await sp.web.lists.getByTitle("KOB").items.select("Id,Title,KOBCode").get();
    return items.map((item: { Id: number; Title: string; KOBCode: string }) => ({
      key: item.Id,
      text: item.Title,
      data: item.KOBCode
    }));
  }
  
  public getGroupStatusMap() {
    return this.groupStatusMap;
  }

  public async getItemById(listName: string, id: number): Promise<any> {
    // First query: plain fields only
    const baseItem = await sp.web.lists.getByTitle(listName).items.getById(id)
      .select(
        "Id,Title, AddressSameAsInProfile_x00a0_, AssetSize, BillingAddress1, BillingAddress2, BillingAddressSameAsInProfile, BillingCity, BillingContactName, BillingEmailAddress, BillingPinCode, BillingSPOC, BillingTelephoneNumber, BusinessAddress1, BusinessAddress2, BusinessAddressPinCode, BusinessAddressSameAsInProfile, BusinessCIty, BusinessContactName, CertifiedLicenseIssuedByRBISigne, CertifiedROCCertificate, CreditInstitutionName, CreditRiskContact, DataAddress1, DataAddress2, DataAddressSameAsInProfile, DataCity, DataContactName, DataEmailAddress, DataPinCode, DataTelephoneNumber, DateOfMembership, Devation_x002f_ExpectionApproval, DuplicateName_x002f_Code_x002f_S, EmailAddress, GSTNumberExists, ITSPOC,  LandlineNumberOffice, LatestBalancesheet_x002f_AnnualR, LetterApplicationProspectsLetter, LetterAuthorityIssued, LicenseCancellation_x002f_OtherT, MCAWebsiteChecked, MembershipApplicationForm, MemberShortCode, Mobile, NameOfCoreBanking, NewMemberCode, NewRegisteredNumber, NominatedCity, NominatedNodalAddress1, NominatedNodalAddress2, NominatedNodalEmailAddress, NominatedNodalFaxNumber, NominatedNodalLandlineNumber, NominatedNodalOfficerName, NominatedNodalPincode, OperatingRuleBook, PAN, PaymentAnnualFee, PaymentOfMembershipFee, PricingAnnexure, Rate_x0020_Type, RBI_x002f_RegulatoryListChecked, RegisteredCity, RegisteredOfficeAddress1, RegisteredOfficeAddress2, RegisteredOfficePinCode, Remarks, Senddocumentsby, SiteVerificationreport, TAN, TypeofPricing,  TVRVerification, TVRComment, WebSearchPerformed, MCACheck, RBICheck, OFACCheck, LitigationComment, Status, LegalAction, LegalComment, FTPSupportAction, FTPSupportComment, IDSupportAction, IDSupportComment"
      )
      .get();

    // Second query: lookup/person fields only
    const lookupItem = await sp.web.lists.getByTitle(listName).items.getById(id)
    .select(
      "Author/Id", "Author/Title", "Author/EMail",
      "KAM/Id", "KAM/Title", "KAM/EMail",
      "KOMGrouping/Id", "KOMGrouping/Title",
      "GSTBillingState/Id", "GSTBillingState/Title",
      "NewBillingState/Id", "NewBillingState/Title",
      "NewBusinessAddressState/Id", "NewBusinessAddressState/Title",
      "NewDataState/Id", "NewDataState/Title",
      "NewNominatedNodalState/Id", "NewNominatedNodalState/Title",
      "NewRegisteredOfficeState/Id", "NewRegisteredOfficeState/Title"
    )
    .expand(
      "Author",
      "KAM",
      "KOMGrouping",
      "GSTBillingState",
      "NewBillingState",
      "NewBusinessAddressState",
      "NewDataState",
      "NewNominatedNodalState",
      "NewRegisteredOfficeState"
    )
    .get();

  // Merge results
  return { ...baseItem, ...lookupItem };
}

// public async getItemById(listName: string, id: number): Promise<any> {
//     // First query: plain fields only
//     const baseItem = await sp.web.lists.getByTitle(listName).items.getById(id)
//       .select(
//         "Id,Title, AddressSameAsInProfile_x00a0_, AssetSize, BillingAddress1, BillingAddress2, BillingAddressSameAsInProfile, BillingCity, BillingContactName, BillingEmailAddress, BillingPinCode, BillingSPOC, BillingTelephoneNumber, BusinessAddress1, BusinessAddress2, BusinessAddressPinCode, BusinessAddressSameAsInProfile, BusinessCIty, BusinessContactName, CertifiedLicenseIssuedByRBISigne, CertifiedROCCertificate, CreditInstitutionName, CreditRiskContact, DataAddress1, DataAddress2, DataAddressSameAsInProfile, DataCity, DataContactName, DataEmailAddress, DataPinCode, DataTelephoneNumber, DateOfMembership, Devation_x002f_ExpectionApproval, DuplicateName_x002f_Code_x002f_S, EmailAddress, GSTNumberExists, ITSPOC,  LandlineNumberOffice, LatestBalancesheet_x002f_AnnualR, LetterApplicationProspectsLetter, LetterAuthorityIssued, LicenseCancellation_x002f_OtherT, MCAWebsiteChecked, MembershipApplicationForm, MemberShortCode, Mobile, NameOfCoreBanking, NewMemberCode, NewRegisteredNumber, NominatedCity, NominatedNodalAddress1, NominatedNodalAddress2, NominatedNodalEmailAddress, NominatedNodalFaxNumber, NominatedNodalLandlineNumber, NominatedNodalOfficerName, NominatedNodalPincode, OperatingRuleBook, PAN, PaymentAnnualFee, PaymentOfMembershipFee, PricingAnnexure, Rate_x0020_Type, RBI_x002f_RegulatoryListChecked, RegisteredCity, RegisteredOfficeAddress1, RegisteredOfficeAddress2, RegisteredOfficePinCode, Remarks, Senddocumentsby, SiteVerificationreport, TAN, TypeofPricing,  TVRVerification, TVRComment, WebSearchPerformed, MCACheck, RBICheck, OFACCheck, LitigationComment, Status, LegalAction, LegalComment, FTPSupportAction, FTPSupportComment, IDSupportAction, IDSupportComment"
//       )
//       .get();

//     // Second query: lookup/person fields only
//     const lookupItem = await sp.web.lists.getByTitle(listName).items.getById(id)
//     .select(
//       "Author/Id", "Author/Title", "Author/EMail",
//       "KAM/Id", "KAM/Title", "KAM/EMail",
//       "KOMGroupingId",
//       "GSTBillingStateId",
//       "NewBillingStateId",
//       "NewBusinessAddressStateId",
//       "NewDataStateId",
//       "NewNominatedNodalStateId",
//       "NewRegisteredOfficeStateId"
//     )
//     .expand(
//       "Author",
//       "KAM"
//       // "KOMGrouping",
//       // "GSTBillingState",
//       // "NewBillingState",
//       // "NewBusinessAddressState",
//       // "NewDataState",
//       // "NewNominatedNodalState",
//       // "NewRegisteredOfficeState"
//     )
//     .get();

//   // Merge results
//   return { ...baseItem, ...lookupItem };
// }
}
 