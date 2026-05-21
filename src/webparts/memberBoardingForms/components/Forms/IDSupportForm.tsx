import * as React from "react";
import styles from "../MemberBoardingForms.module.scss";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

import { sp } from '@pnp/sp';

interface IIDSupportFormProps {
  item: any;
}

interface IIDSupportFormState {
  formData: any;
}

export default class IDSupportForm extends React.Component<IIDSupportFormProps, IIDSupportFormState> {
  constructor(props: IIDSupportFormProps) {
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
  //         IDSupportAction: this.state.formData.IDSupportAction,
  //         IDSupportComment: this.state.formData.IDSupportComment,
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

  switch (this.state.formData.IDSupportAction) {
    case "Approve":
      if (this.state.formData.FTPSupportAction === "Approved") {
        newStatus = "Assigned To Clouser";
      } else {
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
      IDSupportAction: this.state.formData.IDSupportAction,
          IDSupportComment: this.state.formData.IDSupportComment,
    };

    // Only include Status if newStatus is not null
    if (newStatus !== null) {
      updatePayload.Status = newStatus;
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
        <h3 className={styles.sectionsLabel}>ID Support Action Form</h3>
        {/* <p>ID: {formData.Id}</p>
        <p>Title: {formData.Title}</p>
        <p>Status: {formData.Status}</p> */}

        <div className={styles.formRow}>
          <Dropdown
            label="IDSupport Action"
            options={[
              { key: 'Approved', text: 'Approved' },
              { key: 'SendBack', text: 'SendBack' },
              { key: 'Reject', text: 'Reject' }
            ]}
            onChanged={(option: IDropdownOption) =>
              this.handleInputChange('IDSupportAction', option.text)
            }
          />
          <TextField
            label="IDSupport Comment"
            multiline
            value={formData.LegalComment || ""}
            onChanged={(newValue) =>
              this.handleInputChange('IDSupportComment', newValue)
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

