import * as React from "react";
import { IStackTokens, PrimaryButton, Stack, TextField } from "@fluentui/react";
import { sp } from "@pnp/sp";
import { ContextStore } from "../Context/ContextStore";
import styles from "../WorkflowBillApproval.module.scss";
import { useNavigate } from "react-router-dom";

export type IFormDetails = {
  location: string;
  plantCode: string;
  startDate: string;
  materialCodes: string;
  remarks: string;
};
// Example formatting
export const stackTokens: IStackTokens = { childrenGap: 40 };

const InitiaterForm: React.FC = () => {
  const { context } = React.useContext(ContextStore);

  const navigate = useNavigate();

  const [formDetails, setFormDetails] = React.useState<IFormDetails>({
    location: "",
    plantCode: "",
    startDate: "",
    materialCodes: "",
    remarks: "",
  });
  const [loading, setLoading] = React.useState(false);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const { name, value } = e.target;
    setFormDetails((prevDetails) => ({
      ...prevDetails,
      [name]: value,
    }));
  };

  const handleSubmit = (e: React.FormEvent<HTMLFormElement>): void => {
    e.preventDefault();

    // Add a row data in the list
    // const siteId = "cdfec0f5-6017-47aa-b95e-bdd953db733f"; // workflow-bill-approval
    const listId = "86892207-d198-453b-9b56-01044bc52533"; // Form Entry

    try {
      setLoading(true);
      (async () => {
        // Create a new list item :contentReference[oaicite:10]{index=10}
        const result = await sp.web.lists.getById(listId).items.add({
          location: formDetails.location, //"Mumbai Office",
          plantCode: formDetails.plantCode, //"PLNT-001",
          startDate: formDetails.startDate, //"2025-06-01",
          remarks: formDetails.remarks, //"Initial entry via PnP Graph",
          redirectURL:
            context.pageContext.web.absoluteUrl +
            context.pageContext.web.serverRelativeUrl +
            "#/",
          currStep: 1,
        });
        console.log("Created item:", result);
        setFormDetails({
          location: "",
          plantCode: "",
          startDate: "",
          materialCodes: "",
          remarks: "",
        });
      })().catch((error) => {
        console.error("catch haha Error creating item:", error);
        navigate("/err500");
      });
    } catch (error) {
      console.error("Error creating item:", error);
    } finally {
      setLoading(false);
    }
  };

  const disableSubmit =
    formDetails.location === "" ||
    formDetails.plantCode === "" ||
    formDetails.startDate === "" ||
    formDetails.remarks === "" ||
    loading;
  return (
    <div style={{ margin: "auto", maxWidth: "600px" }}>
      <form action="" onSubmit={handleSubmit}>
        <Stack horizontal tokens={stackTokens} horizontalAlign="stretch">
          <TextField
            value={formDetails.location}
            onChange={handleChange}
            label="Location"
            name="location"
          />
          <TextField
            value={formDetails.plantCode}
            onChange={handleChange}
            label="Plant Code & Name"
            name="plantCode"
          />
          <TextField
            value={formDetails.startDate}
            onChange={handleChange}
            label="Start Date"
            name="startDate"
          />
        </Stack>
        <>
          {/* <TextField
              value={formDetails.materialCodes}
              onChange={handleChange}
              label="Material Codes (Max: 12)"
              name="materialCodes"
              multiline
              rows={12}
              resizable={false}
            /> */}
        </>
        <TextField
          value={formDetails.remarks}
          onChange={handleChange}
          label="Remarks"
          multiline
          /* rows={12} */ name="remarks"
        />

        <div className={styles.links}>
          <a href="#/gmForm/3">gmForm</a>
        </div>
        <PrimaryButton
          text={loading ? "Submitting..." : "Submit"}
          disabled={disableSubmit}
          type="submit"
          style={{ marginTop: "12px", float: "right" }}
        />
      </form>
    </div>
  );
};

export default InitiaterForm;
