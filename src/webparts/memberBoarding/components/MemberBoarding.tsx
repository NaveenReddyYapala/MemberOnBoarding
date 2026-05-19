import * as React from 'react';
import styles from './MemberBoarding.module.scss';
import { IMemberBoardingProps } from './IMemberBoardingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices from '../components/SpService/spService';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
//import { sp } from '@pnp/sp';
import { Item, sp } from "@pnp/sp/presets/all";
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

export interface ILookupValue {
  Id: number;
  Title: string;
}
export interface IPeoplePickerUser {
  Id: number;
  Title: string;   // Display name
  EMail: string;   // Email address
}

export interface IMemberOnBoardingState {
  //formData: { [key: string]: any };
  formData: {
    [key: string]: any;
    AddressSameAsInProfile: any;
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
  komOptions: IDropdownOption[];
  stateOptions: IDropdownOption[];
  KAMName: any;
  selectedFiles: File[];
  LitigationFileSelection: File[];
}

export default class MemberBoarding extends React.Component<IMemberBoardingProps, IMemberOnBoardingState> {
  private spService: spservices;
  constructor(props: IMemberBoardingProps) {
    super(props);
    this.state = {
      formData: {
        AddressSameAsInProfile: "",
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
      komOptions: [],
      stateOptions: [],
      KAMName: [],
      selectedFiles: [],
      LitigationFileSelection: []
    };
    this.spService = new spservices(this.props.context);
  }
  public async componentDidMount() {
    try {
      //this._loadDropdownData();
      const komOptions = await this.spService.GetKomGroupOptions();
      const stateOptions = await this.spService.GetStateOptions();
      this.setState({ komOptions, stateOptions });

      // Check query string
      const queryParams = new UrlQueryParameterCollection(window.location.href);
      const idParam = queryParams.getValue("itemId");

      if (idParam) {
        const itemId = parseInt(idParam, 10);
        const item = await this.spService.getItemById("Membership OnBoarding Request", itemId);
        // Populate formData with item values
        const mappedFormDate = {
          AddressSameAsInProfile: item.AddressSameAsInProfile,
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
          GSTBillingStateId: item.GSTBillingState? {Id: item.GSTBillingState.Id, Title: item.GSTBillingState.Title}: null,
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
          NewBillingStateId: item.NewBillingState ? {Id: item.NewBillingState.Id, Title: item.NewBillingState.Title}: null,
          NewBusinessAddressStateId: item.NewBusinessAddressState ? {Id: item.NewBusinessAddressState.Id, Title: item.NewBusinessAddressState.Title}: null,
          NewDataStateId: item.NewDataState ? {Id: item.NewDataState.Id, Title: item.NewDataState.Title}: null,
          NewMemberCode: item.NewMemberCode,
          NewNominatedNodalStateId: item.NewNominatedNodalState ? {Id: item.NewNominatedNodalState.Id, Title: item.NewNominatedNodalState.Title}: null,
          NewRegisteredNumber: item.NewRegisteredNumber,
          NewRegisteredOfficeStateId: item.NewRegisteredOfficeState ? {Id: item.NewRegisteredOfficeState.Id, Title: item.NewRegisteredOfficeState.Title}: null,
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
        this.setState({ formData: mappedFormDate });
      }
    } catch (error) {
      console.log("error", error)
    }
  }
  private _submit = () => {
    console.log("Form submitted:", this.state.formData);
  };

  private _cancel = () => {
    this.setState({});
  };
  
  private _saveDraft = async () => {
    try {
      const { formData, KAMName, selectedFiles, LitigationFileSelection } = this.state;
      const queryParams = new UrlQueryParameterCollection(window.location.href);
      const idParam = queryParams.getValue("itemId");

      const payload: any = { Status: "Draft" };

      // Map values correctly
      Object.keys(formData).forEach(key => {
        const value = formData[key];
        if (value !== undefined && value !== null && value !== "") {
          if (typeof value === "boolean") {
            payload[key] = value;
          } else if (value instanceof Date) {
            payload[key] = value.toISOString();
          } else if (value.Id) {
            // Lookup field
            payload[`${key}`] = value.Id;
          } else {
            payload[key] = value;
          }
        }
      });

      if (KAMName && KAMName.length > 0) {
        payload.KAMId = KAMName[0].Id;
      }

      let itemId: number;
      if (idParam) {
        // Update existing item
        itemId = parseInt(idParam, 10);
        await sp.web.lists.getByTitle("Membership OnBoarding Request").items.getById(itemId).update(payload);
      } else {
        // Add new item
        const item = await sp.web.lists.getByTitle("Membership OnBoarding Request").items.add(payload);
        itemId = item.data.ID;
        const title = `MO${this.padLeft(itemId.toString(), 4, '0')}`;
        await sp.web.lists.getByTitle("Membership OnBoarding Request").items.getById(itemId).update({ Title: title });
      }

      // Upload files
      if (selectedFiles.length > 0) {
        await this.uploadFiles(`MO${this.padLeft(itemId.toString(), 4, '0')}`, selectedFiles, "DocumentReviewFile");
      }
      if (LitigationFileSelection.length > 0) {
        await this.uploadFiles(`MO${this.padLeft(itemId.toString(), 4, '0')}`, LitigationFileSelection, "LitigationFile");
      }
    } catch (error) {
      console.log("error", error);
    }
  };

  // private _saveDraft = async () => {
  //   try {
  //     const { formData, KAMName, selectedFiles, LitigationFileSelection } = this.state;
  //     const payload: any = { Status: "Draft" };
  //     Object.keys(formData).forEach(key => {
  //       const value = formData[key];
  //       if (value !== undefined && value !== null && value !== "") {
  //         // Convert checkboxes to boolean
  //         if (typeof value === "boolean") {
  //           payload[key] = value;
  //         } else {
  //           payload[key] = value;
  //         }
  //       }
  //     });
  //     if (KAMName && KAMName.length > 0) {
  //       payload.KAMId = KAMName[0].Id;
  //     }
  //     const item = await sp.web.lists.getByTitle("Membership OnBoarding Request").items.add(payload);
  //     console.log("Draft saved:", item);

  //     const itemId = item.data.ID;
  //     const title = `MO${this.padLeft(itemId.toString(), 4, '0')}`;
  //     await sp.web.lists.getByTitle("Membership OnBoarding Request").items.getById(itemId).update({
  //       Title: title
  //     });

  //     if (selectedFiles.length > 0) {
  //       await this.uploadFiles(title, selectedFiles, "DocumentReviewFile");
  //     }

  //     if (LitigationFileSelection.length > 0) {
  //       await this.uploadFiles(title, LitigationFileSelection, "LitigationFile");
  //     }
  //   } catch (error) {
  //     console.log("error", error);
  //   }
  // };

  private async uploadFiles(title: string, files: File[], team: string) {
    const libraryUrl = "/sites/NaveenReddy/MemberOnBoardingFiles";

    const uploadPromises = files.map(file => {
      const uniqueFileName = `${title}-${file.name}`;
      //const libraryUrl = '/sites/NaveenReddy/MemberOnBoardingFiles';
      return sp.web.getFolderByServerRelativeUrl(libraryUrl)
        .files.add(uniqueFileName, file, true)
        .then(result => result.file.getItem())
        .then((fileItem: any) => {
          return fileItem.update({
            ReqNumber: title,
            FileUplaodTeam: team
          });
        });
    });
    await Promise.all(uploadPromises);
  }
  private padLeft(value: string, width: number, paddingChar: string): string {
    const padding = Array(width - value.length + 1).join(paddingChar);
    return padding + value;
  }
  private handleInputChange = (fieldName: string, value?: any) => {
    this.setState(prevState => ({
      formData: {
        ...prevState.formData,
        [fieldName]: value
      }
    }));
  };
  private _onFormatDate = (date: Date): string => {
    // Format the date as '01 Sept 2024'
    const options: Intl.DateTimeFormatOptions = { day: '2-digit', month: 'short', year: 'numeric' };
    const formattedDate = date.toLocaleDateString('en-US', options);
    return formattedDate;
  };
  private onPeoplePickerChange = (items: any[]) => {
    if (items.length > 0) {
      const selectedUsers: IPeoplePickerUser[] = items.map(user => ({
        Id: user.id,
        EMail: user.secondaryText,
        Title: user.text
      }));
      this.setState({ KAMName: selectedUsers });
    } else {
      this.setState({ KAMName: [] });
    }
  };

  private handleFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files: File[] = [];
    if (event.target.files) {
      for (let i = 0; i < event.target.files.length; i++) {
        files.push(event.target.files[i]);
      }
    }
    this.setState({ selectedFiles: files });
  }
  private handleLitigationFileSelection = (event: React.ChangeEvent<HTMLInputElement>) => {
    const LitigationFiles: File[] = [];
    if (event.target.files) {
      for (let i = 0; i < event.target.files.length; i++) {
        LitigationFiles.push(event.target.files[i]);
      }
    }
    this.setState({ LitigationFileSelection: LitigationFiles });
  };

  public render(): React.ReactElement<IMemberBoardingProps> {
    return (
      <div className={styles.memberBoarding} >

        {/* Section 1 */}
        <h3 className={styles.sectionsLabel}> Institution Details</h3>
        <div className={styles.formRow}>
          <TextField
            label="Name of Institution"
            value={this.state.formData.CreditInstitutionName || ""}
            onChanged={(newValue) => this.handleInputChange('CreditInstitutionName', newValue)}
          />
          <Dropdown
            label="Type of Institution"
            options={this.state.komOptions}
            selectedKey={this.state.formData.KOMGroupingId ? this.state.formData.KOMGroupingId.Id : undefined}
            onChanged={(option: IDropdownOption) => {
              this.handleInputChange('KOMGroupingId', { Id: option.key, Title: option.text });
              this.handleInputChange('NewMemberCode', option.data as string);
            }}
          />

          <TextField
            label="Member Code"
            value={this.state.formData.NewMemberCode || ""}
            onChanged={(newValue) => this.handleInputChange('NewMemberCode', newValue)}
          />
          <TextField
            label="Member Short Name"
            value={this.state.formData.MemberShortCode || ""}
            onChanged={(newValue) => this.handleInputChange('MemberShortCode', newValue)}
          />
          <Dropdown
            label="GST Customer Type"
            selectedKey={this.state.formData.GSTNumberExists ? this.state.formData.GSTNumberExists : undefined}
            options={[
              { key: 'Registered', text: 'Registered' },
              { key: 'Unregistered', text: 'Unregistered' },
              { key: 'Export', text: 'Export' },
              { key: 'Deemed Export', text: 'Deemed Export' },
              { key: 'Not liable to Register', text: 'Not liable to Register' },
              { key: 'Exempted', text: 'Exempted' }
            ]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('GSTNumberExists', option.text)}
          />
          {(
            this.state.formData.GSTNumberExists === "Registered" ||
            this.state.formData.GSTNumberExists === "Export" ||
            this.state.formData.GSTNumberExists === "Deemed Export"
          ) && (
              <TextField
                label="Registration Number"
                value={this.state.formData.NewRegisteredNumber || ""}
                onChanged={(newValue) => this.handleInputChange('NewRegisteredNumber', newValue)}
              />
            )}
          <Dropdown
            label="GST Billing State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.GSTBillingStateId ? this.state.formData.GSTBillingStateId.Id : undefined}
            onChanged={(option: IDropdownOption) => this.handleInputChange('GSTBillingStateId', option.key)}
          />
          {/* <PeoplePicker
            context={this.props.context}
            titleText="KAM"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            ensureUser={true}
            selectedItems={this.onPeoplePickerChange}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          /> */}
          <PeoplePicker
            context={this.props.context}
            titleText="KAM"
            personSelectionLimit={1}
            ensureUser={true}
            selectedItems={this.onPeoplePickerChange}
            defaultSelectedUsers={this.state.formData.KAMId ? [this.state.formData.KAMId.Title] : []}
            principalTypes={[PrincipalType.User]}
          />


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
            onChanged={(option: IDropdownOption) => this.handleInputChange('TypeofPricing', option.text)}
          />
          <DatePicker
            label="Date Of Membership"
            value={this.state.formData.DateOfMembership || undefined}
            onSelectDate={(date) => this.handleInputChange('DateOfMembership', date)}
            formatDate={this._onFormatDate}
          />
          {(
            this.state.formData.GSTNumberExists === "Unregistered" ||
            this.state.formData.GSTNumberExists === "Not liable to Register" ||
            this.state.formData.GSTNumberExists === "Exempted"
          ) && (
              <TextField
                label="Remarks"
                value={this.state.formData.Remarks || ""}
                multiline rows={3}
                onChanged={(newValue) => this.handleInputChange('Remarks', newValue)}
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
            onChanged={(newValue) => this.handleInputChange('RegisteredOfficeAddress1', newValue)}
          />
          <TextField
            label="Address 2"
            multiline
            value={this.state.formData.RegisteredOfficeAddress2 || ""}
            onChanged={(newValue) => this.handleInputChange('RegisteredOfficeAddress2', newValue)}
          />
          <TextField
            label="City"
            value={this.state.formData.RegisteredCity || ""}
            onChanged={(newValue) => this.handleInputChange('RegisteredCity', newValue)}
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewRegisteredOfficeStateId ? this.state.formData.NewRegisteredOfficeStateId.Id : undefined}
            onChanged={(option: IDropdownOption) => this.handleInputChange('NewRegisteredOfficeStateId', option.key)}
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.RegisteredOfficePinCode || ""}
            onChanged={(newValue) => this.handleInputChange('RegisteredOfficePinCode', newValue)}
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
            onChanged={(option: IDropdownOption) => this.handleInputChange('BusinessAddressSameAsInProfile', option.text)}
          />
          <TextField
            label="City"
            value={this.state.formData.BusinessCIty || ""}
            onChanged={(newValue) => this.handleInputChange('BusinessCIty', newValue)}
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewBusinessAddressStateId ? this.state.formData.NewBusinessAddressStateId.Id : undefined}
            onChanged={(option: IDropdownOption) => this.handleInputChange('NewBusinessAddressStateId', option.key)}
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.BusinessAddressPinCode || ""}
            onChanged={(newValue) => this.handleInputChange('BusinessAddressPinCode', newValue)}
          />
          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.BusinessAddress1 || ""}
            onChanged={(newValue) => this.handleInputChange('BusinessAddress1', newValue)}
          />
          <TextField
            label="Address 2"
            multiline
            value={this.state.formData.BusinessAddress2 || ""}
            onChanged={(newValue) => this.handleInputChange('BusinessAddress2', newValue)}
          />

        </div>

        {/* Section 4 */}
        <h3 className={styles.sectionsLabel}> Other Details</h3>
        <div className={styles.formRow}>
          <TextField
            label="Landline Number Office"
            value={this.state.formData.LandlineNumberOffice || ""}
            onChanged={(newValue) => this.handleInputChange('LandlineNumberOffice', newValue)}
          />
          <TextField
            label="Mobile"
            value={this.state.formData.Mobile || ""}
            onChanged={(newValue) => this.handleInputChange('Mobile', newValue)}
          />
          <TextField
            label="EmailAddress"
            value={this.state.formData.EmailAddress || ""}
            onChanged={(newValue) => this.handleInputChange('EmailAddress', newValue)}
          />

          <TextField
            label="PAN"
            value={this.state.formData.PAN || ""}
            onChanged={(newValue) => this.handleInputChange('PAN', newValue)}
          />
          <TextField
            label="Public Static IP address"
            value={this.state.formData.TAN || ""}
            onChanged={(newValue) => this.handleInputChange('TAN', newValue)}
          />
          <TextField
            label="Cheque details"
            value={this.state.formData.AssetSize || ""}
            onChanged={(newValue) => this.handleInputChange('AssetSize', newValue)}
          />
          <Dropdown
            label="Send Documents By"
            options={[
              { key: 'Mail', text: 'Mail' },
              { key: 'Courier', text: 'Courier' }
            ]}
            selectedKey={this.state.formData.Senddocumentsby ? this.state.formData.Senddocumentsby : undefined}
            onChanged={(option: IDropdownOption) => {
              this.handleInputChange('Senddocumentsby', option.text);
            }}
          />

          <TextField
            label="Rate Type"
            value={this.state.formData.Rate_x0020_Type || ""}
            onChanged={(newValue) => this.handleInputChange('Rate_x0020_Type', newValue)}
          />
          <TextField
            label="Billing SPOC"
            value={this.state.formData.BillingSPOC || ""}
            onChanged={(newValue) => this.handleInputChange('BillingSPOC', newValue)}
          />
          <TextField
            label="IT SPOC"
            value={this.state.formData.ITSPOC || ""}
            onChanged={(newValue) => this.handleInputChange('ITSPOC', newValue)}
          />
        </div>

        {/* Section 5*/}
        <h3 className={styles.sectionsLabel}> Contact Details</h3>
        <div className={styles.formRow}>
          <TextField
            label="Nominated Nodal Officer Name"
            value={this.state.formData.NominatedNodalOfficerName || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalOfficerName', newValue)}
          />
          <TextField
            label="Credit Risk Contact"
            value={this.state.formData.CreditRiskContact || ""}
            onChanged={(newValue) => this.handleInputChange('CreditRiskContact', newValue)}
          />
          <TextField
            label="Business Contact Name"
            value={this.state.formData.BusinessContactName || ""}
            onChanged={(newValue) => this.handleInputChange('BusinessContactName', newValue)}
          />
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.AddressSameAsInProfile ? this.state.formData.AddressSameAsInProfile : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('AddressSameAsInProfile', option.text)}
          />
          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.NominatedNodalAddress1 || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalAddress1', newValue)}
          />
          <TextField label="Address 2" multiline
            value={this.state.formData.NominatedNodalAddress2 || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalAddress2', newValue)}
          />
          <TextField label="City"
            value={this.state.formData.NominatedCity || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedCity', newValue)}
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewNominatedNodalStateId ? this.state.formData.NewNominatedNodalStateId.Id : undefined}
            onChanged={(option: IDropdownOption) => this.handleInputChange('NewNominatedNodalStateId', option.key)}
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.NominatedNodalPincode || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalPincode', newValue)}
          />
          <TextField
            label="Landline Number"
            value={this.state.formData.NominatedNodalLandlineNumber || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalLandlineNumber', newValue)}
          />
          <TextField
            label="FaxNumber"
            value={this.state.formData.NominatedNodalFaxNumber || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalFaxNumber', newValue)}
          />
          <TextField
            label="NominatedNodalEmailAddress"
            value={this.state.formData.NominatedNodalEmailAddress || ""}
            onChanged={(newValue) => this.handleInputChange('NominatedNodalEmailAddress', newValue)}
          />
        </div>

        {/* Section 6*/}
        <h3 className={styles.sectionsLabel}>Alternate Nodal Officer</h3>
        <div className={styles.formRow}>
          <TextField label="Data Contact Name"
            value={this.state.formData.DataContactName || ""}
            onChanged={(newValue) => this.handleInputChange('DataContactName', newValue)}
          />
          <TextField label="Name Of Core Banking"
            value={this.state.formData.NameOfCoreBanking || ""}
            onChanged={(newValue) => this.handleInputChange('NameOfCoreBanking', newValue)}
          />
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.DataAddressSameAsInProfile ? this.state.formData.DataAddressSameAsInProfile : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('DataAddressSameAsInProfile', option.text)}
          />

          <TextField label="City"
            value={this.state.formData.DataCity || ""}
            onChanged={(newValue) => this.handleInputChange('DataCity', newValue)}
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewDataStateId ? this.state.formData.NewDataStateId.Id : undefined}
            onChanged={(option: IDropdownOption) => this.handleInputChange('NewDataStateId', option.key)}
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.DataPinCode || ""}
            onChanged={(newValue) => this.handleInputChange('DataPinCode', newValue)}
          />
          <TextField
            label="Landline Number"
            value={this.state.formData.DataTelephoneNumber || ""}
            onChanged={(newValue) => this.handleInputChange('DataTelephoneNumber', newValue)}
          />
          <TextField
            label="DataEmailAddress"
            value={this.state.formData.DataEmailAddress || ""}
            onChanged={(newValue) => this.handleInputChange('DataEmailAddress', newValue)}
          />

          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.DataAddress1 || ""}
            onChanged={(newValue) => this.handleInputChange('DataAddress1', newValue)}
          />
          <TextField label="Address 2" multiline
            value={this.state.formData.DataAddress2 || ""}
            onChanged={(newValue) => this.handleInputChange('DataAddress2', newValue)}
          />
        </div>

        {/* Section 7*/}
        <h3 className={styles.sectionsLabel}>Billing Contact (Optional)</h3>
        <div className={styles.formRow}>
          <TextField label="Billing Contact Name"
            value={this.state.formData.BillingContactName || ""}
            onChanged={(newValue) => this.handleInputChange('BillingContactName', newValue)}
          />
          <Dropdown
            label="Address Same As In Profile"
            selectedKey={this.state.formData.BillingAddressSameAsInProfile ? this.state.formData.BillingAddressSameAsInProfile : undefined}
            options={[
              { key: 'Yes', text: 'Yes' },
              { key: 'No', text: 'No' }
            ]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('BillingAddressSameAsInProfile', option.text)}
          />
          <TextField
            label="Address 1"
            multiline
            value={this.state.formData.BillingAddress1 || ""}
            onChanged={(newValue) => this.handleInputChange('BillingAddress1', newValue)}
          />
          <TextField label="Address 2" multiline
            value={this.state.formData.BillingAddress2 || ""}
            onChanged={(newValue) => this.handleInputChange('BillingAddress2', newValue)}
          />
          <TextField label="City"
            value={this.state.formData.BillingCity || ""}
            onChanged={(newValue) => this.handleInputChange('BillingCity', newValue)}
          />
          <Dropdown
            label="State"
            options={this.state.stateOptions}
            selectedKey={this.state.formData.NewBillingStateId ? this.state.formData.NewBillingStateId.Id : undefined}
            onChanged={(option: IDropdownOption) => this.handleInputChange('NewBillingStateId', option.key)}
          />
          <TextField
            label="Pin Code"
            value={this.state.formData.BillingPinCode || ""}
            onChanged={(newValue) => this.handleInputChange('BillingPinCode', newValue)}
          />
          <TextField
            label="Landline Number"
            value={this.state.formData.BillingTelephoneNumber || ""}
            onChanged={(newValue) => this.handleInputChange('BillingTelephoneNumber', newValue)}
          />
          <TextField
            label="BillingEmailAddress"
            value={this.state.formData.BillingEmailAddress || ""}
            onChanged={(newValue) => this.handleInputChange('BillingEmailAddress', newValue)}
          />
        </div>

        {/* Section 8: Checkboxes */}
        <h3 className={styles.sectionsLabel}>CheckList</h3>
        <div className={styles.formRow}>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.MembershipApplicationForm}
              onChange={e => this.handleInputChange("MembershipApplicationForm", e.target.checked)}
            />
            Membership Application Form
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.PricingAnnexure}
              onChange={e => this.handleInputChange("PricingAnnexure", e.target.checked)}
            />
            Pricing Annexure
          </label>
          {/* Repeat for all checkboxes */}
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LetterApplicationProspectsLetter}
              onChange={e => this.handleInputChange("LetterApplicationProspectsLetter", e.target.checked)}
            />
            Letter of Application on Prospects letter head
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LetterAuthorityIssued}
              onChange={e => this.handleInputChange("LetterAuthorityIssued", e.target.checked)}
            />
            Copy of Letter of Authority issued by Member
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.PaymentAnnualFee}
              onChange={e => this.handleInputChange("PaymentAnnualFee", e.target.checked)}
            />
            Payment of Annual Fee
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.PaymentOfMembershipFee}
              onChange={e => this.handleInputChange("PaymentOfMembershipFee", e.target.checked)}
            />
            Payment Of Membership Fee
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.OperatingRuleBook}
              onChange={e => this.handleInputChange("OperatingRuleBook", e.target.checked)}
            />
            Operating Rule Book
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.Devation_x002f_ExpectionApproval}
              onChange={e => this.handleInputChange("Devation_x002f_ExpectionApproval", e.target.checked)}
            />
            Any deviation or exception approval taken
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.CertifiedROCCertificate}
              onChange={e => this.handleInputChange("CertifiedROCCertificate", e.target.checked)}
            />
            Certified true copy of the  ROC Certificate, Financials
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.CertifiedLicenseIssuedByRBISigne}
              onChange={e => this.handleInputChange("CertifiedLicenseIssuedByRBISigne", e.target.checked)}
            />
            <span>Certified (Signed & Stamped) true copy of the license issued by RBI/Regulatory Authority</span>
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.SiteVerificationreport}
              onChange={e => this.handleInputChange("SiteVerificationreport", e.target.checked)}
            />
            Site Verification report
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LatestBalancesheet_x002f_AnnualR}
              onChange={e => this.handleInputChange("LatestBalancesheet_x002f_AnnualR", e.target.checked)}
            />
            Latest Balance sheet / Annual Report
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.DuplicateName_x002f_Code_x002f_S}
              onChange={e => this.handleInputChange("DuplicateName_x002f_Code_x002f_S", e.target.checked)}
            />
            Duplicate Name/Code Short Name Checked In Member Code Creation File
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.LicenseCancellation_x002f_OtherT}
              onChange={e => this.handleInputChange("LicenseCancellation_x002f_OtherT", e.target.checked)}
            />
            License Cancellation / Other Termination List Checked
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.MCAWebsiteChecked}
              onChange={e => this.handleInputChange("MCAWebsiteChecked", e.target.checked)}
            />
            MCA Website Checked
          </label>
          <label>
            <input
              type="checkbox"
              checked={!!this.state.formData.RBI_x002f_RegulatoryListChecked}
              onChange={e => this.handleInputChange("RBI_x002f_RegulatoryListChecked", e.target.checked)}
            />
            RBI / Regulatory List Checked
          </label>
          <label>
            <div className={styles.uplaodLabel}>Uplaod Document Review Files</div>
            <input id="fileInput"
              type="file" multiple
              onChange={this.handleFileSelection} />
          </label>
        </div>
        <div>
          {(this.state.selectedFiles.length > 0) &&
            <div>
              <h1>Attachments :</h1>
              {/* <div className={styles.attachBox}> */}
              {this.state.selectedFiles && this.state.selectedFiles.length > 0 &&
                <div>
                  <div>
                    <ul >
                      {this.state.selectedFiles.map((file, index) => (
                        <li key={index} >
                          <span>{file.name}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              }

            </div>
          }
        </div>

        {/* Section 9: TVR Verification */}
        <h3 className={styles.sectionsLabel}> TVR Verification</h3>
        <div className={styles.formRow}>
          <Dropdown label="TVR Verification"
          selectedKey={this.state.formData.TVRVerification ? this.state.formData.TVRVerification : undefined}
            options={[{ key: 'Connect', text: 'Connect' }, { key: 'Non-connect', text: 'Non-connect' }, { key: 'Approved', text: 'Approved' }, { key: 'Reject', text: 'Reject' }]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('TVRVerification', option.text)}
          />
          <TextField label="Comment" multiline
            value={this.state.formData.TVRComment || ""}
            onChanged={(newValue) => this.handleInputChange('TVRComment', newValue)}
          />
        </div>

        {/* Section 10: Litigation Checks */}
        <h3 className={styles.sectionsLabel}> Litigation Checks</h3>
        <div className={styles.formRow}>
          <Dropdown label="Web search performed"
          selectedKey={this.state.formData.WebSearchPerformed ? this.state.formData.WebSearchPerformed : undefined}
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('WebSearchPerformed', option.text)}
          />
          <Dropdown label="MCA Check"
          selectedKey={this.state.formData.MCACheck ? this.state.formData.MCACheck : undefined}
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('MCACheck', option.text)}
          />
          <Dropdown label="RBI website Check"
          selectedKey={this.state.formData.RBICheck ? this.state.formData.RBICheck : undefined}
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('RBICheck', option.text)}
          />
          <Dropdown label="OFAC / Sanction check"
          selectedKey={this.state.formData.OFACCheck ? this.state.formData.OFACCheck : undefined}
            options={[{ key: 'Yes', text: 'Yes' }, { key: 'No', text: 'No' }, { key: 'N/A', text: 'N/A' }]}
            onChanged={(option: IDropdownOption) => this.handleInputChange('OFACCheck', option.text)}
          />
          <TextField label="Litigation Comment" multiline
            value={this.state.formData.LitigationComment || ""}
            onChanged={(newValue) => this.handleInputChange('LitigationComment', newValue)}
          />
        </div>
        <div>
          <label>
            <div className={styles.uplaodLabel}>Upload Litigation Files</div>
            <input id="fileInput"
              type="file" multiple
              onChange={this.handleLitigationFileSelection} />
          </label>
        </div>
        <div>
          {(this.state.LitigationFileSelection.length > 0) &&
            <div>
              <h1>Attachments :</h1>
              {/* <div className={styles.attachBox}> */}
              {this.state.LitigationFileSelection && this.state.LitigationFileSelection.length > 0 &&
                <div>
                  <div>
                    <ul >
                      {this.state.LitigationFileSelection.map((file, index) => (
                        <li key={index} >
                          <span>{file.name}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              }
            </div>
          }
        </div>

        {/* Action buttons */}
        <div>
          <div style={{ marginTop: "20px" }}>
            <button type="button" onClick={this._saveDraft}>Save as Draft</button>
            <button type="button" onClick={this._submit}>Submit</button>
            <button type="button" onClick={this._cancel}>Cancel</button>
          </div>
        </div>

      </div >
    );
  }
}
