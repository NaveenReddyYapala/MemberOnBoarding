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
import { Item } from '@pnp/sp/items';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';

interface IMemberOnBoardingFile {
  Id: number;
  ReqNumber: string;
  FileUplaodTeam: string;
  FileLeafRef: string;
  FileRef: string;
}

export interface ILookupValue {
  Id: number;
  Title: string;
}
export interface IPeoplePickerUser {
  Id: number;
  Title: string;   // Display name
  EMail: string;   // Email address
}

export interface IMemberBoardingFormsState {
  formData: {
    [key: string]: any;
    AddressSameAsInProfile_x00a0_: any;
    AssetSize: any;
    BillingAddress1: any;
    BillingAddress2: any;
    BillingAddressSameAsInProfile: any;
    BillingCity: any;
    BillingContactName: any;
    BillingEmailAddress: any;
    BillingPinCode: any;
    BillingSPOC: any;
    BillingTelephoneNumber: any;
    BusinessAddress1: any;
    BusinessAddress2: any;
    BusinessAddressPinCode: any;
    BusinessAddressSameAsInProfile: any;
    BusinessCIty: any;
    BusinessContactName: any;
    CertifiedLicenseIssuedByRBISigne: boolean | null;
    CertifiedROCCertificate: boolean | null;
    CreditInstitutionName: any;
    CreditRiskContact: any;
    DataAddress1: any;
    DataAddress2: any;
    DataAddressSameAsInProfile: any;
    DataCity: any;
    DataContactName: any;
    DataEmailAddress: any;
    DataPinCode: any;
    DataTelephoneNumber: any;
    DateOfMembership: Date | null;
    Devation_x002f_ExpectionApproval: boolean | null;
    DuplicateName_x002f_Code_x002f_S: boolean | null;
    EmailAddress: any;
    GSTBillingStateId: ILookupValue | null;
    GSTNumberExists: any;
    ITSPOC: any;
    KAMId: IPeoplePickerUser | null;
    KOMGroupingId: ILookupValue | null;
    LandlineNumberOffice: any;
    LatestBalancesheet_x002f_AnnualR: boolean | null;
    LetterApplicationProspectsLetter: boolean | null;
    LetterAuthorityIssued: boolean | null;
    LicenseCancellation_x002f_OtherT: boolean | null;
    MCAWebsiteChecked: boolean | null;
    MembershipApplicationForm: boolean | null;
    MemberShortCode: any;
    Mobile: any;
    NameOfCoreBanking: any;
    NewBillingStateId: ILookupValue | null;
    NewBusinessAddressStateId: ILookupValue | null;
    NewDataStateId: ILookupValue | null;
    NewMemberCode: any;
    NewNominatedNodalStateId: ILookupValue | null;
    NewRegisteredNumber: any;
    NewRegisteredOfficeStateId: ILookupValue | null;
    NominatedCity: any;
    NominatedNodalAddress1: any;
    NominatedNodalAddress2: any;
    NominatedNodalEmailAddress: any;
    NominatedNodalFaxNumber: any;
    NominatedNodalLandlineNumber: any;
    NominatedNodalOfficerName: any;
    NominatedNodalPincode: any;
    OperatingRuleBook: boolean | null;
    PAN: any;
    PaymentAnnualFee: boolean | null;
    PaymentOfMembershipFee: boolean | null;
    PricingAnnexure: boolean | null;
    Rate_x0020_Type: any;
    RBI_x002f_RegulatoryListChecked: boolean | null;
    RegisteredCity: any;
    RegisteredOfficeAddress1: any;
    RegisteredOfficeAddress2: any;
    RegisteredOfficePinCode: any;
    Remarks: any;
    Senddocumentsby: any;
    SiteVerificationreport: boolean | null;
    TAN: any;
    TypeofPricing: any;

    TVRVerification: any;
    TVRComment: any;
    WebSearchPerformed: any;
    MCACheck: any;
    RBICheck: any;
    OFACCheck: any;
    LitigationComment: any;
    Status: any;
  };
  userGroups: string[];
  getDocumentReviewFiles: IMemberOnBoardingFile[];
  getLitigationFiles: IMemberOnBoardingFile[];
  komOptions: IDropdownOption[];
  stateOptions: IDropdownOption[];
}

export default class MemberBoardingForms extends React.Component<IMemberBoardingFormsProps, IMemberBoardingFormsState> {
  private spService: spservices;
  constructor(props: IMemberBoardingFormsProps) {
    super(props);
    this.state = {
      getDocumentReviewFiles: [],
      getLitigationFiles: [],
      userGroups: [],
      komOptions: [],
      stateOptions: [],
      formData: {
        AddressSameAsInProfile_x00a0_: "",
        AssetSize: "",
        BillingAddress1: "",
        BillingAddress2: "",
        BillingAddressSameAsInProfile: "",
        BillingCity: "",
        BillingContactName: "",
        BillingEmailAddress: "",
        BillingPinCode: "",
        BillingSPOC: "",
        BillingTelephoneNumber: "",
        BusinessAddress1: "",
        BusinessAddress2: "",
        BusinessAddressPinCode: "",
        BusinessAddressSameAsInProfile: "",
        BusinessCIty: "",
        BusinessContactName: "",
        CertifiedLicenseIssuedByRBISigne: null,
        CertifiedROCCertificate: null,
        CreditInstitutionName: "",
        CreditRiskContact: "",
        DataAddress1: "",
        DataAddress2: "",
        DataAddressSameAsInProfile: "",
        DataCity: "",
        DataContactName: "",
        DataEmailAddress: "",
        DataPinCode: "",
        DataTelephoneNumber: "",
        DateOfMembership: null,
        Devation_x002f_ExpectionApproval: null,
        DuplicateName_x002f_Code_x002f_S: null,
        EmailAddress: "",
        GSTBillingStateId: null,
        GSTNumberExists: "",
        ITSPOC: "",
        KAMId: null,
        KOMGroupingId: null,
        LandlineNumberOffice: "",
        LatestBalancesheet_x002f_AnnualR: null,
        LetterApplicationProspectsLetter: null,
        LetterAuthorityIssued: null,
        LicenseCancellation_x002f_OtherT: null,
        MCAWebsiteChecked: null,
        MembershipApplicationForm: null,
        MemberShortCode: "",
        Mobile: "",
        NameOfCoreBanking: "",
        NewBillingStateId: null,
        NewBusinessAddressStateId: null,
        NewDataStateId: null,
        NewMemberCode: "",
        NewNominatedNodalStateId: null,
        NewRegisteredNumber: "",
        NewRegisteredOfficeStateId: null,
        NominatedCity: "",
        NominatedNodalAddress1: "",
        NominatedNodalAddress2: "",
        NominatedNodalEmailAddress: "",
        NominatedNodalFaxNumber: "",
        NominatedNodalLandlineNumber: "",
        NominatedNodalOfficerName: "",
        NominatedNodalPincode: "",
        OperatingRuleBook: null,
        PAN: "",
        PaymentAnnualFee: null,
        PaymentOfMembershipFee: null,
        PricingAnnexure: null,
        Rate_x0020_Type: "",
        RBI_x002f_RegulatoryListChecked: null,
        RegisteredCity: "",
        RegisteredOfficeAddress1: "",
        RegisteredOfficeAddress2: "",
        RegisteredOfficePinCode: "",
        Remarks: "",
        Senddocumentsby: "",
        SiteVerificationreport: null,
        TAN: "",
        TypeofPricing: "",

        TVRVerification: "",
        TVRComment: "",
        WebSearchPerformed: "",
        MCACheck: "",
        RBICheck: "",
        OFACCheck: "",
        LitigationComment: "",
        Status: ""
      },
    };
    this.spService = new spservices(this.props.context);
  }
  public async componentDidMount() {
    try {
      const style = document.createElement('style');
      style.innerHTML = `    
           .ms-Image-image { display: none; }
           .ms-Persona-imageArea{ display: none; }
           .ms-Persona-imageArea.imageArea-227 { display: none; }
            `;
      document.head.appendChild(style);
      const komOptions = await this.spService.GetKomGroupOptions();
      const stateOptions = await this.spService.GetStateOptions();
      this.setState({ komOptions, stateOptions });

      const queryParams = new UrlQueryParameterCollection(window.location.href);
      const idParam = queryParams.getValue("ItemId");

      let item: any | null = null;

      if (idParam) {
        const itemId = parseInt(idParam, 10);
        item = await this.spService.getItemById("Membership On-Boarding Request", itemId);
      }
      const mappedFormData = {
        AddressSameAsInProfile_x00a0_: item.AddressSameAsInProfile_x00a0_,
        AssetSize: item.AssetSize,
        BillingAddress1: item.BillingAddress1,
        BillingAddress2: item.BillingAddress2,
        BillingAddressSameAsInProfile: item.BillingAddressSameAsInProfile,
        BillingCity: item.BillingCity,
        BillingContactName: item.BillingContactName,
        BillingEmailAddress: item.BillingEmailAddress,
        BillingPinCode: item.BillingPinCode,
        BillingSPOC: item.BillingSPOC,
        BillingTelephoneNumber: item.BillingTelephoneNumber,
        BusinessAddress1: item.BusinessAddress1,
        BusinessAddress2: item.BusinessAddress2,
        BusinessAddressPinCode: item.BusinessAddressPinCode,
        BusinessAddressSameAsInProfile: item.BusinessAddressSameAsInProfile,
        BusinessCIty: item.BusinessCIty,
        BusinessContactName: item.BusinessContactName,
        CertifiedLicenseIssuedByRBISigne: !!item.CertifiedLicenseIssuedByRBISigne,
        CertifiedROCCertificate: !!item.CertifiedROCCertificate,
        CreditInstitutionName: item.CreditInstitutionName,
        CreditRiskContact: item.CreditRiskContact,
        DataAddress1: item.DataAddress1,
        DataAddress2: item.DataAddress2,
        DataAddressSameAsInProfile: item.DataAddressSameAsInProfile,
        DataCity: item.DataCity,
        DataContactName: item.DataContactName,
        DataEmailAddress: item.DataEmailAddress,
        DataPinCode: item.DataPinCode,
        DataTelephoneNumber: item.DataTelephoneNumber,
        DateOfMembership: item.DateOfMembership ? new Date(item.DateOfMembership) : null,
        Devation_x002f_ExpectionApproval: !!item.Devation_x002f_ExpectionApproval,
        DuplicateName_x002f_Code_x002f_S: !!item.DuplicateName_x002f_Code_x002f_S,
        EmailAddress: item.EmailAddress,
        GSTBillingStateId: item.GSTBillingState ? { Id: item.GSTBillingState.Id, Title: item.GSTBillingState.Title } : null,
        GSTNumberExists: item.GSTNumberExists,
        ITSPOC: item.ITSPOC,
        KAMId: item.KAM ? { Id: item.KAM.Id, Title: item.KAM.Title, EMail: item.KAM.EMail } : null,
        KOMGroupingId: item.KOMGrouping
          ? { Id: item.KOMGrouping.Id, Title: item.KOMGrouping.Title }
          : null,
        LandlineNumberOffice: item.LandlineNumberOffice,
        LatestBalancesheet_x002f_AnnualR: !!item.LatestBalancesheet_x002f_AnnualR,
        LetterApplicationProspectsLetter: !!item.LetterApplicationProspectsLetter,
        LetterAuthorityIssued: !!item.LetterAuthorityIssued,
        LicenseCancellation_x002f_OtherT: !!item.LicenseCancellation_x002f_OtherT,
        MCAWebsiteChecked: !!item.MCAWebsiteChecked,
        MembershipApplicationForm: !!item.MembershipApplicationForm,
        MemberShortCode: item.MemberShortCode,
        Mobile: item.Mobile,
        NameOfCoreBanking: item.NameOfCoreBanking,
        NewBillingStateId: item.NewBillingState ? { Id: item.NewBillingState.Id, Title: item.NewBillingState.Title } : null,
        NewBusinessAddressStateId: item.NewBusinessAddressState ? { Id: item.NewBusinessAddressState.Id, Title: item.NewBusinessAddressState.Title } : null,
        NewDataStateId: item.NewDataState ? { Id: item.NewDataState.Id, Title: item.NewDataState.Title } : null,
        NewMemberCode: item.NewMemberCode,
        NewNominatedNodalStateId: item.NewNominatedNodalState ? { Id: item.NewNominatedNodalState.Id, Title: item.NewNominatedNodalState.Title } : null,
        NewRegisteredNumber: item.NewRegisteredNumber,
        NewRegisteredOfficeStateId: item.NewRegisteredOfficeState ? { Id: item.NewRegisteredOfficeState.Id, Title: item.NewRegisteredOfficeState.Title } : null,
        NominatedCity: item.NominatedCity,
        NominatedNodalAddress1: item.NominatedNodalAddress1,
        NominatedNodalAddress2: item.NominatedNodalAddress2,
        NominatedNodalEmailAddress: item.NominatedNodalEmailAddress,
        NominatedNodalFaxNumber: item.NominatedNodalFaxNumber,
        NominatedNodalLandlineNumber: item.NominatedNodalLandlineNumber,
        NominatedNodalOfficerName: item.NominatedNodalOfficerName,
        NominatedNodalPincode: item.NominatedNodalPincode,
        OperatingRuleBook: !!item.OperatingRuleBook,
        PAN: item.PAN,
        PaymentAnnualFee: !!item.PaymentAnnualFee,
        PaymentOfMembershipFee: !!item.PaymentOfMembershipFee,
        PricingAnnexure: !!item.PricingAnnexure,
        Rate_x0020_Type: item.Rate_x0020_Type,
        RBI_x002f_RegulatoryListChecked: !!item.RBI_x002f_RegulatoryListChecked,
        RegisteredCity: item.RegisteredCity,
        RegisteredOfficeAddress1: item.RegisteredOfficeAddress1,
        RegisteredOfficeAddress2: item.RegisteredOfficeAddress2,
        RegisteredOfficePinCode: item.RegisteredOfficePinCode,
        Remarks: item.Remarks,
        Senddocumentsby: item.Senddocumentsby,
        SiteVerificationreport: !!item.SiteVerificationreport,
        TAN: item.TAN,
        TypeofPricing: item.TypeofPricing,

        TVRVerification: item.TVRVerification,
        TVRComment: item.TVRComment,
        WebSearchPerformed: item.WebSearchPerformed,
        MCACheck: item.MCACheck,
        RBICheck: item.RBICheck,
        OFACCheck: item.OFACCheck,
        LitigationComment: item.LitigationComment,
        Status: item.Status
      }
      const userGroups = await this.spService.getUserGroups();
      const getFiles: IMemberOnBoardingFile[] = await sp.web.lists
        .getByTitle("MemberOnBoardingFiles")
        .items.filter(`ReqNumber eq '${item.Title}'`)
        .select("Id,ReqNumber,FileUplaodTeam,FileLeafRef,FileRef")
        .get();

      const getDocumentReviewFiles = getFiles.filter(f => f.FileUplaodTeam === "DocumentReviewFile");
      const getLitigationFiles = getFiles.filter(f => f.FileUplaodTeam === "LitigationFile");
      this.setState({ formData: mappedFormData, userGroups, getDocumentReviewFiles, getLitigationFiles });
    } catch (error) {
      console.log("error", error);
    }
  }


  private renderForm() {
    const { formData, userGroups } = this.state;
    if (!formData) return null;

    const status = formData.Status;
    const map = this.spService.getGroupStatusMap();

    // Check each group → status mapping
    if (userGroups.indexOf("LegalTeam") > -1 && status === map["LegalTeam"]) {
      return <LegalForm item={formData} />;
    }

    if (userGroups.indexOf("CheckerTeam") > -1 && status === map["CheckerTeam"]) {
      return <CheckerForm item={formData} />;
    }
    if (userGroups.indexOf("Member On-Boarding Finance Team") > -1 && status === map["Member On-Boarding Finance Team"]) {
      return <FinanceForm item={formData} />;
    }
    if (userGroups.indexOf("ID Support Team") > -1 && status === map["ID Support Team"]) {
      return <IDSupportForm item={formData} />;
    }
    if (userGroups.indexOf("FTP Support Team") > -1 && status === map["FTP Support Team"]) {
      return <FTPSupportForm item={formData} />;
    }
    if (userGroups.indexOf("Membership On-Boarding Closure") > -1 && status === map["Membership On-Boarding Closure"]) {
      return <MOClouserForm item={formData} />;
    }

    return <div>No form available for your group/status.</div>;
  }

  public render(): React.ReactElement<IMemberBoardingFormsProps> {
    return (
      <div className={styles.memberBoardingForms} >

        {/* Section 1 */}
        <h3 className={styles.sectionsLabel}> Institution Details</h3>
        <div className={styles.formRow}>
          <TextField
            label="Name of Institution"
            value={this.state.formData.CreditInstitutionName || ""}
            disabled
          />
          <Dropdown
            label="Type of Institution"
            options={this.state.komOptions}
            selectedKey={this.state.formData.KOMGroupingId ? this.state.formData.KOMGroupingId.Id : undefined}
            disabled
          />

          <TextField
            label="Member Code"
            value={this.state.formData.NewMemberCode || ""}
            disabled
          />
          <TextField
            label="Member Short Name"
            value={this.state.formData.MemberShortCode || ""}
            disabled
          />
          <Dropdown
            label="GST Customer Type"
            options={[
              { key: 'Registered', text: 'Registered' },
              { key: 'Unregistered', text: 'Unregistered' },
              { key: 'Export', text: 'Export' },
              { key: 'Deemed Export', text: 'Deemed Export' },
              { key: 'Not liable to Register', text: 'Not liable to Register' },
              { key: 'Exempted', text: 'Exempted' }
            ]}
            selectedKey={this.state.formData.GSTNumberExists ? this.state.formData.GSTNumberExists : undefined}
            disabled
          />
          {(
            this.state.formData.GSTNumberExists === "Registered" ||
            this.state.formData.GSTNumberExists === "Export" ||
            this.state.formData.GSTNumberExists === "Deemed Export"
          ) && (
              <TextField
                label="Registration Number"
                value={this.state.formData.NewRegisteredNumber || ""}
                disabled
              />
            )}
          <Dropdown
            label="GST Billing State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.GSTBillingStateId ? this.state.formData.GSTBillingStateId.Id : undefined}
            disabled
          />
          <TextField
            label="KAM"
            value={this.state.formData.KAMId ? this.state.formData.KAMId.Title : ""}
            disabled
          />
          {/* <PeoplePicker
            context={this.props.context}
            titleText="KAM"
            personSelectionLimit={1}
            ensureUser={true}
            disabled
            defaultSelectedUsers={this.state.formData.KAMId ? [this.state.formData.KAMId.Title] : []}
          /> */}

          <Dropdown
            label="Type Of Pricing"
            selectedKey={this.state.formData.TypeofPricing ? this.state.formData.TypeofPricing : undefined}
            options={[
              { key: 'Financial Institutions', text: 'Financial Institutions' },
              { key: 'Private / MNC', text: 'Private / MNC' },
              { key: 'PSU', text: 'PSU' },
              { key: 'MFI', text: 'MFI' },
              { key: 'Insurance', text: 'Insurance' },
              { key: 'Telecom', text: 'Telecom' },
              { key: 'Limited to Financial Institutions', text: 'Limited to Financial Institutions' },
              { key: 'Limited to Private / MNC', text: 'Limited to Private / MNC' },
              { key: 'Limited to PSU', text: 'Limited to PSU' },
              { key: 'Limited to MFI', text: 'Limited to MFI' },
              { key: 'Limited to Insurance', text: 'Limited to Insurance' },
              { key: 'Limited to Telecom', text: 'Limited to Telecom' },
              { key: 'Not Applicable', text: 'Not Applicable' }
            ]}
            disabled
          />
          <DatePicker
            label="Date Of Membership"
            value={this.state.formData.DateOfMembership || undefined}
            disabled
          />
          {(
            this.state.formData.GSTNumberExists === "Unregistered" ||
            this.state.formData.GSTNumberExists === "Not liable to Register" ||
            this.state.formData.GSTNumberExists === "Exempted"
          ) && (
              <TextField
                label="Remarks"
                value={this.state.formData.Remarks || ""}
                multiline
                disabled
              />
            )}
        </div>

        {/* Section 2 */}
        <h3 className={styles.sectionsLabel}> Registered Office Address</h3>
        <div className={styles.formRow}>
          <TextField
            label="Address 1"
            value={this.state.formData.RegisteredOfficeAddress1 || ""}
            multiline
            disabled
          />
          <TextField
            label="Address 2"
            multiline
            value={this.state.formData.RegisteredOfficeAddress2 || ""}
            disabled
          />
          <TextField
            label="City"
            value={this.state.formData.RegisteredCity || ""}
            disabled
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewRegisteredOfficeStateId ? this.state.formData.NewRegisteredOfficeStateId.Id : undefined}
            disabled
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.RegisteredOfficePinCode || ""}
            disabled
          />
        </div>

        {/* Section 3 */}
        <h3 className={styles.sectionsLabel}> Business Address (HO)</h3>
        <div className={styles.formRow}>
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.BusinessAddressSameAsInProfile ? this.state.formData.BusinessAddressSameAsInProfile : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            disabled
          />
          <TextField
            label="City"
            value={this.state.formData.BusinessCIty || ""}
            disabled
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewBusinessAddressStateId ? this.state.formData.NewBusinessAddressStateId.Id : undefined}
            disabled
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.BusinessAddressPinCode || ""}
            disabled
          />
          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.BusinessAddress1 || ""}
            disabled
          />
          <TextField
            label="Address 2"
            multiline
            value={this.state.formData.BusinessAddress2 || ""}
            disabled
          />

        </div>

        {/* Section 4 */}
        <h3 className={styles.sectionsLabel}> Other Details</h3>
        <div className={styles.formRow}>
          <TextField
            label="Landline Number Office"
            value={this.state.formData.LandlineNumberOffice || ""}
            disabled
          />
          <TextField
            label="Mobile"
            value={this.state.formData.Mobile || ""}
            disabled
          />
          <TextField
            label="EmailAddress"
            value={this.state.formData.EmailAddress || ""}
            disabled
          />

          <TextField
            label="PAN"
            value={this.state.formData.PAN || ""}
            disabled
          />
          <TextField
            label="Public Static IP address"
            value={this.state.formData.TAN || ""}
            disabled
          />
          <TextField
            label="Cheque details"
            value={this.state.formData.AssetSize || ""}
            disabled
          />
          <Dropdown
            label="Send Documents By"
            selectedKey={this.state.formData.Senddocumentsby ? this.state.formData.Senddocumentsby : undefined}
            options={[
              { key: 'Mail', text: 'Mail' },
              { key: 'Courier', text: 'Courier' }
            ]}
            disabled
          />

          <TextField
            label="Rate Type"
            value={this.state.formData.Rate_x0020_Type || ""}
            disabled
          />
          <TextField
            label="Billing SPOC"
            value={this.state.formData.BillingSPOC || ""}
            disabled
          />
          <TextField
            label="IT SPOC"
            value={this.state.formData.ITSPOC || ""}
            disabled
          />
        </div>

        {/* Section 5*/}
        <h3 className={styles.sectionsLabel}> Contact Details</h3>
        <div className={styles.formRow}>
          <TextField
            label="Nominated Nodal Officer Name"
            value={this.state.formData.NominatedNodalOfficerName || ""}
            disabled
          />
          <TextField
            label="Credit Risk Contact"
            value={this.state.formData.CreditRiskContact || ""}
            disabled
          />
          <TextField
            label="Business Contact Name"
            value={this.state.formData.BusinessContactName || ""}
            disabled
          />
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.AddressSameAsInProfile_x00a0_ ? this.state.formData.AddressSameAsInProfile_x00a0_ : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            disabled
          />
          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.NominatedNodalAddress1 || ""}
            disabled
          />
          <TextField label="Address 2" multiline
            value={this.state.formData.NominatedNodalAddress2 || ""}
            disabled
          />
          <TextField label="City"
            value={this.state.formData.NominatedCity || ""}
            disabled
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewNominatedNodalStateId ? this.state.formData.NewNominatedNodalStateId.Id : undefined}
            disabled
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.NominatedNodalPincode || ""}
            disabled
          />
          <TextField
            label="Landline Number"
            value={this.state.formData.NominatedNodalLandlineNumber || ""}
            disabled
          />
          <TextField
            label="FaxNumber"
            value={this.state.formData.NominatedNodalFaxNumber || ""}
            disabled
          />
          <TextField
            label="NominatedNodalEmailAddress"
            value={this.state.formData.NominatedNodalEmailAddress || ""}
            disabled
          />
        </div>

        {/* Section 6*/}
        <h3 className={styles.sectionsLabel}>Alternate Nodal Officer</h3>
        <div className={styles.formRow}>
          <TextField label="Data Contact Name"
            value={this.state.formData.DataContactName || ""}
            disabled
          />
          <TextField label="Name Of Core Banking"
            value={this.state.formData.NameOfCoreBanking || ""}
            disabled
          />
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.DataAddressSameAsInProfile ? this.state.formData.DataAddressSameAsInProfile : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            disabled
          />

          <TextField label="City"
            value={this.state.formData.DataCity || ""}
            disabled
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewDataStateId ? this.state.formData.NewDataStateId.Id : undefined}
            disabled
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.DataPinCode || ""}
            disabled
          />
          <TextField
            label="Landline Number"
            value={this.state.formData.DataTelephoneNumber || ""}
            disabled
          />
          <TextField
            label="DataEmailAddress"
            value={this.state.formData.DataEmailAddress || ""}
            disabled
          />

          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.DataAddress1 || ""}
            disabled
          />
          <TextField label="Address 2" multiline
            value={this.state.formData.DataAddress2 || ""}
            disabled
          />
        </div>

        {/* Section 7*/}
        <h3 className={styles.sectionsLabel}>Billing Contact (Optional)</h3>
        <div className={styles.formRow}>
          <TextField label="Billing Contact Name"
            value={this.state.formData.BillingContactName || ""}
            disabled
          />
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.BillingAddressSameAsInProfile ? this.state.formData.BillingAddressSameAsInProfile : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            disabled
          />
          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.BillingAddress1 || ""}
            disabled
          />
          <TextField label="Address 2" multiline
            value={this.state.formData.BillingAddress2 || ""}
            disabled
          />
          <TextField label="City"
            value={this.state.formData.BillingCity || ""}
            disabled
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewBillingStateId ? this.state.formData.NewBillingStateId.Id : undefined}
            disabled
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.BillingPinCode || ""}
            disabled
          />
          <TextField
            label="Landline Number"
            value={this.state.formData.BillingTelephoneNumber || ""}
            disabled
          />
          <TextField
            label="BillingEmailAddress"
            value={this.state.formData.BillingEmailAddress || ""}
            disabled
          />
        </div>

        {/* Section 8: Checkboxes */}
        <h3 className={styles.sectionsLabel}>CheckList</h3>
        <div className={styles.formRow}>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.MembershipApplicationForm}
              disabled
            />
            Membership Application Form
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.PricingAnnexure}
              disabled
            />
            Pricing Annexure
          </label>
          {/* Repeat for all checkboxes */}
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LetterApplicationProspectsLetter}
              disabled
            />
            Letter of Application on Prospects letter head
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LetterAuthorityIssued}
              disabled
            />
            Copy of Letter of Authority issued by Member
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.PaymentAnnualFee}
              disabled
            />
            Payment of Annual Fee
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.PaymentOfMembershipFee}
              disabled
            />
            Payment Of Membership Fee
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.OperatingRuleBook}
              disabled
            />
            Operating Rule Book
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.Devation_x002f_ExpectionApproval}
              disabled
            />
            Any deviation or exception approval taken
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.CertifiedROCCertificate}
              disabled
            />
            Certified true copy of the  ROC Certificate, Financials
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.CertifiedLicenseIssuedByRBISigne}
              disabled
            />
            <span>Certified (Signed & Stamped) true copy of the license issued by RBI/Regulatory Authority</span>
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.SiteVerificationreport}
              disabled
            />
            Site Verification report
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LatestBalancesheet_x002f_AnnualR}
              disabled
            />
            Latest Balance sheet / Annual Report
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.DuplicateName_x002f_Code_x002f_S}
              disabled
            />
            Duplicate Name/Code Short Name Checked In Member Code Creation File
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LicenseCancellation_x002f_OtherT}
              disabled
            />
            License Cancellation / Other Termination List Checked
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.MCAWebsiteChecked}
              disabled
            />
            MCA Website Checked
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.RBI_x002f_RegulatoryListChecked}
              disabled
            />
            RBI / Regulatory List Checked
          </label>
        </div>
        {(this.state.getDocumentReviewFiles.length > 0) &&
          <div>
            <h2 className={styles.fileHeaders}>Document review Files</h2>
            <div className={styles.attachmentsContainer}>
              {this.state.getDocumentReviewFiles.map(file => (
                <div key={file.Id} className={styles.fileCardView}>
                  <a href={file.FileRef} className={styles.fileNameView} target="_blank" rel="noopener noreferrer">
                    {file.FileLeafRef}
                  </a>
                </div>
              ))}
            </div>
          </div>
        }

        {/* Section 9: TVR Verification */}
        <h3 className={styles.sectionsLabel}> TVR Verification</h3>
        <div className={styles.formRow}>
          <Dropdown label="TVR Verification"
            options={[{ key: 'Connect', text: 'Connect' }, { key: 'Non-connect', text: 'Non-connect' }, { key: 'Approved', text: 'Approved' }, { key: 'Reject', text: 'Reject' }]}
            selectedKey={this.state.formData.TVRVerification ? this.state.formData.TVRVerification : undefined}
            disabled
          />
          <TextField label="Comment" multiline
            value={this.state.formData.TVRComment || ""}
            disabled
          />
        </div>

        {/* Section 10: Litigation Checks */}
        <h3 className={styles.sectionsLabel}> Litigation Checks</h3>
        <div className={styles.formRow}>
          <Dropdown label="Web search performed"
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            selectedKey={this.state.formData.WebSearchPerformed ? this.state.formData.WebSearchPerformed : undefined}
            disabled
          />
          <Dropdown label="MCA Check"
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            selectedKey={this.state.formData.MCACheck ? this.state.formData.MCACheck : undefined}
            disabled
          />
          <Dropdown label="RBI website Check"
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            selectedKey={this.state.formData.RBICheck ? this.state.formData.RBICheck : undefined}
            disabled
          />
          <Dropdown label="OFAC / Sanction check"
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            selectedKey={this.state.formData.OFACCheck ? this.state.formData.OFACCheck : undefined}
            disabled
          />
          <TextField label="Litigation Comment" multiline
            value={this.state.formData.LitigationComment || ""}
            disabled
          />
        </div>


        {(this.state.getLitigationFiles.length > 0) &&
          <div>
            <h2 className={styles.fileHeaders}>Litigation Files</h2>
            <div className={styles.attachmentsContainer}>
              {this.state.getLitigationFiles.map(file => (
                <div key={file.Id} className={styles.fileCardView}>
                  <a href={file.FileRef} className={styles.fileNameView} target="_blank" rel="noopener noreferrer">
                    {file.FileLeafRef}
                  </a>
                </div>
              ))}
            </div>
          </div>
        }

        <div>
          {this.renderForm()}
        </div>
      </div >
    );
  }
}
