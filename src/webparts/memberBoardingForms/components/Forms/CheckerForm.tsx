import * as React from "react";
import styles from "../MemberBoardingForms.module.scss";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

import { sp } from '@pnp/sp';

interface ICheckerFormProps {
  item: any;
}

interface ICheckerFormState {
  formData: any;
}

export default class CheckerForm extends React.Component<ICheckerFormProps, ICheckerFormState> {
  constructor(props: ICheckerFormProps) {
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
        newStatus = "Assigned To Finance";
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
          LegalAction: this.state.formData.LegalAction,
          LegalComment: this.state.formData.LegalComment,
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
        <h3>Checker Action Form</h3>
        {/* <p>ID: {formData.Id}</p>
        <p>Title: {formData.Title}</p>
        <p>Status: {formData.Status}</p> */}

        <div className={styles.formRow}>
          <Dropdown
            label="Checker Action"
            options={[
              { key: 'Approve', text: 'Approve' },
              { key: 'SendBack', text: 'SendBack' },
              { key: 'Reject', text: 'Reject' }
            ]}
            onChanged={(option: IDropdownOption) =>
              this.handleInputChange('CheckerAction', option.text)
            }
          />
          <TextField
            label="Checker Comment"
            multiline
            value={formData.LegalComment || ""}
            onChanged={(newValue) =>
              this.handleInputChange('CheckerComment', newValue)
            }
          />
        </div>
        <div>
          <button type="button" onClick={this.handleSubmit} >Submit</button>
        </div>
      </div>
    );
  }
}

