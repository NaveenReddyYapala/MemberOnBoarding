import * as React from "react";
import styles from "../MemberBoardingForms.module.scss";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

import { sp } from '@pnp/sp';

interface IMOClouserFormProps {
  item: any;
}

interface IMOClouserFormState {
  formData: any;
}

export default class MOClouserForm extends React.Component<IMOClouserFormProps, IMOClouserFormState> {
  constructor(props: IMOClouserFormProps) {
    super(props);
    this.state = { formData: props.item };
  }

  private handleInputChange = (field: string, value: any) => {
    this.setState(prevState => ({
      formData: { ...prevState.formData, [field]: value }
    }));
  };

  private handleSubmit = async () => {
    let newStatus = "";
    switch (this.state.formData.LegalAction) {
      case "Approve":
        newStatus = "Completed";
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
    const idParam = queryParams.getValue("ItemId"); // or "itemId" depending on your URL

    if (!idParam) {
      alert("No ItemId found in query string");
      return;
    }

    const itemId = parseInt(idParam, 10);

    try {
      await sp.web.lists.getByTitle("Membership On-Boarding Request")
        .items.getById(itemId)
        .update({
          MOBClouserAction: this.state.formData.MOBClouserAction,
          MOBClouserComment: this.state.formData.MOBClouserComment,
          MOBClouserEndDate: new Date(),
          Status: newStatus
        });
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
        <h3 className={styles.sectionsLabel}>MOB Clouser Action Form</h3>
        {/* <p>ID: {formData.Id}</p>
        <p>Title: {formData.Title}</p>
        <p>Status: {formData.Status}</p> */}

        <div className={styles.formRow}>
          <Dropdown
            label="MOClouser Action"
            options={[
              { key: 'Approve', text: 'Approve' },
              { key: 'SendBack', text: 'SendBack' },
              { key: 'Reject', text: 'Reject' }
            ]}
            onChanged={(option: IDropdownOption) =>
              this.handleInputChange('MOBClouserAction', option.text)
            }
          />
          <TextField
            label="MOClouser Comment"
            multiline
            value={formData.LegalComment || ""}
            onChanged={(newValue) =>
              this.handleInputChange('MOBClouserComment', newValue)
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

