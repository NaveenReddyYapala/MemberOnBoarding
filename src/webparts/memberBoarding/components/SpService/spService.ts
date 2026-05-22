import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web, SiteGroups, Item } from '@pnp/sp/presets/all';
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export default class spservices {
    constructor(private context: any) {
        sp.setup({ sp: { baseUrl: this.context.pageContext.web.absoluteUrl } });
    }
    public GetItems(): Promise<any[]> {
        return sp.web.lists
            .getByTitle("Membership On-Boarding Request")
            .items
            .select("Id,Title").orderBy("ID", false)
            .top(4999)
            .get();
    }

    // Get dropdown options from KomGroupList
  public async GetKomGroupOptions(): Promise<IDropdownOption[]> {
    const items = await sp.web.lists.getByTitle("KOB").items.select("Id,Title,KOBCode").get();
    return items.map((item: { Id: number; Title: string; KOBCode: string }) => ({
      key: item.Id,
      text: item.Title,
      data: item.KOBCode
    }));
  }
 public async GetNextMemberCode(kobCode: string): Promise<string> {
  // Step 1: Get all member codes starting with kobCode
  const web = Web("https://wisdombat.transunion.com/sites/apps");

const items = await web.lists.getByTitle("Member Master")
  //const items = await sp.web.lists.getByTitle("Member Master")
    .items
    .select("MemberCode")
    .filter(`startswith(MemberCode,'${kobCode}')`)
    .orderBy("MemberCode", true) // ascending
    .get();

  if (!items || items.length === 0) {
    // No existing codes → start from 001
    return kobCode + "0001";
  }

  // Step 2: Get the last (highest) code
  const lastCode = items[items.length - 1].MemberCode; // e.g. "CU100"

  // Step 3: Extract numeric part
  const numPart = parseInt(lastCode.replace(kobCode, ""), 10);

  // Step 4: Increment
  const nextNum = numPart + 1;

  // Step 5: Format with leading zeros (manual padding)
  const nextCode = kobCode + ("0000" + nextNum).slice(-4);

  return nextCode;
}

  // Get dropdown options from StateList
  public async GetStateOptions(): Promise<IDropdownOption[]> {
//     const web = Web("https://wisdombat.transunion.com/sites/apps");

// const items = await web.lists.getByTitle("GST Master")
    const items = await sp.web.lists.getByTitle("MOBStates")
    .items.select("Id, Title, GSTBillingCode").get();
    return items.map((item: { Id: number; Title: string; GSTBillingCode: string }) => ({
      key: item.Id,
      text: item.GSTBillingCode
    }));
  }

  // public async getItemById(listName: string, id: any): Promise<any> {
  //   const item = await sp.web.lists.getByTitle(listName).items.getById(id)
  //   .select("Id", "Title", "AddressSameAsInProfile", "AssetSize", "BillingAddress1", "BillingAddress2", "BillingAddressSameAsInProfile", "BillingCity", "BillingContactName", "BillingEmailAddress", "BillingPinCode", "BillingSPOC", "BillingTelephoneNumber", "BusinessAddress1", "BusinessAddress2", "BusinessAddressPinCode", "BusinessAddressSameAsInProfile", "BusinessCIty", "BusinessContactName", "CertifiedLicenseIssuedByRBISigne", "CertifiedROCCertificate", "CreditInstitutionName", "CreditRiskContact", "DataAddress1", "DataAddress2", "DataAddressSameAsInProfile", "DataCity", "DataContactName", "DataEmailAddress", "DataPinCode", "DataTelephoneNumber", "DateOfMembership", "Devation_x002f_ExpectionApproval", "DuplicateName_x002f_Code_x002f_S", "EmailAddress", "GSTBillingState/Id", "GSTBillingState/Title", "GSTNumberExists", "ITSPOC", "KAMId/Id", "KAMId/Title", "KAMId/EMail", "KOMGrouping/Id", "KOMGrouping/Title", "LandlineNumberOffice", "LatestBalancesheet_x002f_AnnualR", "LetterApplicationProspectsLetter", "LetterAuthorityIssued", "LicenseCancellation_x002f_OtherT", "MCAWebsiteChecked", "MembershipApplicationForm", "MemberShortCode", "Mobile", "NameOfCoreBanking", "NewBillingState/Id", "NewBillingState/Title", "NewBusinessAddressState/Id", "NewBusinessAddressState/Title", "NewDataState/Id", "NewDataState/Title", "NewMemberCode", "NewNominatedNodalState/Id", "NewNominatedNodalState/Title", "NewRegisteredNumber", "NewRegisteredOfficeState/Id", "NewRegisteredOfficeState/Title", "NominatedCity", "NominatedNodalAddress1", "NominatedNodalAddress2", "NominatedNodalEmailAddress", "NominatedNodalFaxNumber", "NominatedNodalLandlineNumber", "NominatedNodalOfficerName", "NominatedNodalPincode", "OperatingRuleBook", "PAN", "PaymentAnnualFee", "PaymentOfMembershipFee", "PricingAnnexure", "Rate_x0020_Type", "RBI_x002f_RegulatoryListChecked", "RegisteredCity", "RegisteredOfficeAddress1", "RegisteredOfficeAddress2", "RegisteredOfficePinCode", "Remarks", "Senddocumentsby", "SiteVerificationreport", "TAN", "TypeofPricing",  "TVRVerification", "TVRComment", "WebSearchPerformed", "MCACheck", "RBICheck", "OFACCheck", "LitigationComment", "Status", "Author/Id", "Author/Title", "Author/EMail")
  //     .expand("Author", "GSTBillingState", "KAM", "KOMGrouping", "NewBillingState", "NewBusinessAddressState", "NewDataState", "NewNominatedNodalState", "NewRegisteredOfficeState")
  //     .get();
  //   return item;
  // }
  public async getItemById(listName: string, id: number): Promise<any> {
    // First query: plain fields only
    const baseItem = await sp.web.lists.getByTitle(listName).items.getById(id)
      .select(
        "Id,Title, AddressSameAsInProfile_x00a0_, AssetSize, BillingAddress1, BillingAddress2, BillingAddressSameAsInProfile, BillingCity, BillingContactName, BillingEmailAddress, BillingPinCode, BillingSPOC, BillingTelephoneNumber, BusinessAddress1, BusinessAddress2, BusinessAddressPinCode, BusinessAddressSameAsInProfile, BusinessCIty, BusinessContactName, CertifiedLicenseIssuedByRBISigne, CertifiedROCCertificate, CreditInstitutionName, CreditRiskContact, DataAddress1, DataAddress2, DataAddressSameAsInProfile, DataCity, DataContactName, DataEmailAddress, DataPinCode, DataTelephoneNumber, DateOfMembership, Devation_x002f_ExpectionApproval, DuplicateName_x002f_Code_x002f_S, EmailAddress, GSTNumberExists, ITSPOC,  LandlineNumberOffice, LatestBalancesheet_x002f_AnnualR, LetterApplicationProspectsLetter, LetterAuthorityIssued, LicenseCancellation_x002f_OtherT, MCAWebsiteChecked, MembershipApplicationForm, MemberShortCode, Mobile, NameOfCoreBanking, NewMemberCode, NewRegisteredNumber, NominatedCity, NominatedNodalAddress1, NominatedNodalAddress2, NominatedNodalEmailAddress, NominatedNodalFaxNumber, NominatedNodalLandlineNumber, NominatedNodalOfficerName, NominatedNodalPincode, OperatingRuleBook, PAN, PaymentAnnualFee, PaymentOfMembershipFee, PricingAnnexure, Rate_x0020_Type, RBI_x002f_RegulatoryListChecked, RegisteredCity, RegisteredOfficeAddress1, RegisteredOfficeAddress2, RegisteredOfficePinCode, Remarks, Senddocumentsby, SiteVerificationreport, TAN, TypeofPricing,  TVRVerification, TVRComment, WebSearchPerformed, MCACheck, RBICheck, OFACCheck, LitigationComment, Status"
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
      "NewRegisteredOfficeState/Id", "NewRegisteredOfficeState/Title",
      "CheckerAction", "CheckerComment", "CheckerEndDate", "CheckerStartDate", "FinanceAction", "FinanceComment", "FinanceEndDate", "FinanceStartDate", "FTPEndDate", "FTPStartDate", "FTPSupportAction", "FTPSupportComment", "IDEndDate", "IDStartDate", "IDSupportAction", "IDSupportComment", "LegalAction", "LegalComment", "LegalEndDate", "LegalStartDate", "MOBClouserAction", "MOBClouserComment", "MOBClouserEndDate", "MOBClouserStartDate"
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


}