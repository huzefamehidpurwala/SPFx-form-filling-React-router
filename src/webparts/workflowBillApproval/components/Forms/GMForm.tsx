import {
  DefaultButton,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";
import { sp } from "@pnp/sp";
import * as React from "react";
import { useNavigate, useParams } from "react-router-dom";
import { IFormDetails, stackTokens } from "./InitiaterForm";
import styles from "../WorkflowBillApproval.module.scss";

const GMForm: React.FC = () => {
  const { formId } = useParams();
  const navigate = useNavigate();

  const [formDetails, setFormDetails] = React.useState<IFormDetails>({
    location: "",
    plantCode: "",
    startDate: "",
    materialCodes: "",
    remarks: "",
  });
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    if (!formId) {
      navigate("/err404");
      return;
    }

    // Fetch the form data
    (async () => {
      // const siteId = "cdfec0f5-6017-47aa-b95e-bdd953db733f"; // workflow-bill-approval
      const listId = "86892207-d198-453b-9b56-01044bc52533"; // Form Entry

      try {
        (async () => {
          // Create a new list item :contentReference[oaicite:10]{index=10}
          const result = await sp.web.lists
            .getById(listId)
            .items.getById(Number(formId))
            .get();
          //   console.log("Created item:", result);
          setFormDetails(result);
        })().catch((error) => {
          console.error("catch haha Error creating item:", error);
          navigate("/err500");
        });
      } catch (error) {
        console.error("Error creating item:", error);
        navigate("/err500");
      } finally {
        setLoading(false);
      }
    })().catch((error) => {
      console.error("catch external creating item:", error);
      navigate("/err500");
    });
  }, [formId]);

  return (
    <div style={{ margin: "auto", maxWidth: "600px" }}>
      {loading ? (
        <p>Fetching details...</p>
      ) : (
        <>
          <Stack horizontal tokens={stackTokens} horizontalAlign="stretch">
            <TextField
              value={formDetails.location}
              disabled
              label="Location"
              name="location"
            />
            <TextField
              value={formDetails.plantCode}
              disabled
              label="Plant Code & Name"
              name="plantCode"
            />
            <TextField
              value={formDetails.startDate}
              disabled
              label="Start Date"
              name="startDate"
            />
          </Stack>
          <TextField
            value={formDetails.remarks}
            disabled
            label="Remarks"
            multiline
            name="remarks"
          />
        </>
      )}
      <div className={styles.links}>
        <a href="#/form">initiaterForm</a>
      </div>
      <Stack horizontal tokens={{ childrenGap: 8 }} style={{ float: "right" }}>
        <DefaultButton text="Reject" onClick={() => {}} />
        <PrimaryButton text="Approve" onClick={() => {}} />
      </Stack>
    </div>
  );
};
export default GMForm;
