import * as React from "react";
import styles from "../MemberBoardingForms.module.scss";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

import { sp } from '@pnp/sp';

interface IFTPSupportFormProps {
  item: any;
}

interface IFTPSupportFormState {
  formData: any;
}

export default class FTPSupportForm extends React.Component<IFTPSupportFormProps, IFTPSupportFormState> {
  constructor(props: IFTPSupportFormProps) {
    super(props);
    this.state = { formData: props.item };
  }

  private handleInputChange = (field: string, value: any) => {
    this.setState(prevState => ({
      formData: { ...prevState.formData, [field]: value }
    }));
  };

  // private handleSubmit = async () => {
  //   let newStatus = "";
  //   switch (this.state.formData.LegalAction) {
  //     case "Approve":
  //       newStatus = "Assigned To Closure";
  //       break;
  //     case "SendBack":
  //       newStatus = "Assigned to Maker";
  //       break;
  //     case "Reject":
  //       newStatus = "Rejected";
  //       break;
  //     default:
  //       newStatus = this.state.formData.Status;
  //   }

  //   const queryParams = new UrlQueryParameterCollection(window.location.href);
  //   const idParam = queryParams.getValue("ItemId"); // or "itemId" depending on your URL

  //   if (!idParam) {
  //     alert("No ItemId found in query string");
  //     return;
  //   }

  //   const itemId = parseInt(idParam, 10);

  //   try {
  //     await sp.web.lists.getByTitle("Membership On-Boarding Request")
  //       .items.getById(itemId)
  //       .update({
  //         FTPSupportAction: this.state.formData.FTPSupportAction,
  //         FTPSupportComment: this.state.formData.FTPSupportComment,
  //         Status: newStatus
  //       });
  //     alert("Form submitted successfully!");
  //   } catch (err) {
  //     console.error("Error updating item:", err);
  //     alert("Error saving data.");
  //   }
  // };

  private handleSubmit = async () => {
  let newStatus: string | null = null; // use null to indicate "don't update"

  switch (this.state.formData.FTPSupportAction) {
    case "Approve":
      if (this.state.formData.IDSupportAction === "Approved") {
        newStatus = "Assigned To Clouser";
      } else {
        // If IDSupportAction is not Approved, don't update Status
        newStatus = null;
      }
      break;
    case "SendBack":
      newStatus = "Assigned to Maker";
      break;
    case "Reject":
      newStatus = "Rejected";
      break;
    default:
      newStatus = this.state.formData.Status;
  }

  const queryParams = new UrlQueryParameterCollection(window.location.href);
  const idParam = queryParams.getValue("ItemId");

  if (!idParam) {
    alert("No ItemId found in query string");
    return;
  }

  const itemId = parseInt(idParam, 10);

  try {
    const updatePayload: any = {
      FTPSupportAction: this.state.formData.FTPSupportAction,
      FTPSupportComment: this.state.formData.FTPSupportComment,
      FTPEndDate: new Date()
    };

    // Only include Status if newStatus is not null
    if (newStatus !== null) {
      updatePayload.Status = newStatus;
      updatePayload.MOBClouserStartDate = new Date();
    }

    await sp.web.lists.getByTitle("Membership On-Boarding Request")
      .items.getById(itemId)
      .update(updatePayload);

    alert("Form submitted successfully!");
  } catch (err) {
    console.error("Error updating item:", err);
    alert("Error saving data.");
  }
};


  render() {
    const { formData } = this.state;
    return (
      <div className={styles.memberBoardingForms}>
        <h3 className={styles.sectionsLabel}>FTP Support Action Form</h3>
        {/* <p>ID: {formData.Id}</p>
        <p>Title: {formData.Title}</p>
        <p>Status: {formData.Status}</p> */}

        <div className={styles.formRow}>
          <Dropdown
            label="FTPSupport Action"
            options={[
              { key: 'Approved', text: 'Approved' },
              { key: 'SendBack', text: 'SendBack' }
            ]}
            onChanged={(option: IDropdownOption) =>
              this.handleInputChange('FTPSupportAction', option.text)
            }
          />
          <TextField
            label="FTPSupport Comment"
            multiline
            value={formData.LegalComment || ""}
            onChanged={(newValue) =>
              this.handleInputChange('FTPSupportComment', newValue)
            }
          />
        </div>
        <div className={styles.btnContainer}>
          <button type="button" className={styles.btnStyles} onClick={this.handleSubmit} >Submit</button>
          <button type="button" className={styles.btnStyles} >Cancel</button>
        </div>
      </div>
    );
  }
}

