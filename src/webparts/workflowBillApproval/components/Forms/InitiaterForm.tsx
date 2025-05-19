import * as React from "react";
import {
  ComboBox,
  DefaultButton,
  Dialog,
  DialogContent,
  DialogFooter,
  DialogType,
  IComboBoxOption,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";
import { sp } from "@pnp/sp";
import { ContextStore } from "../Context/ContextStore";
// import styles from "../WorkflowBillApproval.module.scss";
import { useNavigate, useParams } from "react-router-dom";

export type IFormDetails = {
  location: string;
  plantCode: string;
  startDate: string;
  materialCodes: string;
  remarks: string;
};
// Example formatting
export const stackTokens: IStackTokens = { childrenGap: 40 };
// const siteId = "cdfec0f5-6017-47aa-b95e-bdd953db733f"; // workflow-bill-approval
const listId = "86892207-d198-453b-9b56-01044bc52533"; // Form Entry

const dialogContentProps = {
  type: DialogType.normal,
  title: "Confirm Rejection",
};

const InitiaterForm: React.FC = () => {
  const { spContext: context } = React.useContext(ContextStore);

  const { formId } = useParams();
  const navigate = useNavigate();

  const [formDetails, setFormDetails] = React.useState<IFormDetails>({
    location: "",
    plantCode: "",
    startDate: "",
    materialCodes: "",
    remarks: "",
  });
  const [comBoxSelectedKey, setComBoxSelectedKey] = React.useState({
    location: "",
    plantCode: "",
  });
  const [loading, setLoading] = React.useState(false);
  const [rejectReason, setRejectReason] = React.useState("");
  const [hideDialog, setHideDialog] = React.useState(true);
  const [currStep, setCurrStep] = React.useState(0);

  const toggleHideDialog = (): void => setHideDialog((p) => !p);

  React.useEffect(() => {
    if (!formId) {
      // navigate("/err/404");
      return;
    }

    // Fetch the form data
    (async () => {
      // const siteId = "cdfec0f5-6017-47aa-b95e-bdd953db733f"; // workflow-bill-approval
      const listId = "86892207-d198-453b-9b56-01044bc52533"; // Form Entry

      try {
        setLoading(true);
        await (async () => {
          // Create a new list item :contentReference[oaicite:10]{index=10}
          const result = await sp.web.lists
            .getById(listId)
            .items.getById(Number(formId))
            .get();
          //   console.log("Created item:", result);
          setFormDetails(result);
          setCurrStep(result.currStep);
          setComBoxSelectedKey({
            location: result.location.toLowerCase(),
            plantCode: result.plantCode.toLowerCase(),
          });
        })().catch((error) => {
          console.error("catch haha Error creating item:", error);
          navigate("/err/500", { replace: true });
        });
      } catch (error) {
        console.error("Error creating item:", error);
        navigate("/err/500", { replace: true });
      } finally {
        setLoading(false);
      }
    })().catch((error) => {
      console.error("catch external creating item:", error);
      navigate("/err/500", { replace: true });
    });
  }, [formId]);

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
    const absoluteUrl = context.pageContext.web.absoluteUrl;
    const pageRelativePath = context.pageContext.site.serverRequestPath.replace(
      context.pageContext.site.serverRelativeUrl,
      ""
    );
    const hashRoute = "#/form/";

    (async () => {
      try {
        setLoading(true);
        // Create a new list item :contentReference[oaicite:10]{index=10}
        if (!formId) {
          /* const result = */ await sp.web.lists.getById(listId).items.add({
            location: formDetails.location, //"Mumbai Office",
            plantCode: formDetails.plantCode, //"PLNT-001",
            startDate: formDetails.startDate, //"2025-06-01",
            remarks: formDetails.remarks, //"Initial entry via PnP Graph",
            redirectURL: absoluteUrl + pageRelativePath + hashRoute,
            currStep: 1,
            reasonOfRejection: null,
            rejectedBy: null,
            rejectedFromStep: null,
          });
          navigate("/", { replace: true });
          return;
        }
        /* const result = */ await sp.web.lists
          .getById(listId)
          .items.getById(Number(formId))
          .update({
            location: formDetails.location, //"Mumbai Office",
            plantCode: formDetails.plantCode, //"PLNT-001",
            startDate: formDetails.startDate, //"2025-06-01",
            remarks: formDetails.remarks, //"Initial entry via PnP Graph",
            redirectURL: absoluteUrl + pageRelativePath + hashRoute,
            currStep: 1,
            reasonOfRejection: null,
            rejectedBy: null,
            rejectedFromStep: null,
          });
        navigate("/", { replace: true });
        // console.log("Created item:", result);
        // setComBoxSelectedKey({ location: "", plantCode: "" });
        // setFormDetails({
        //   location: "",
        //   plantCode: "",
        //   startDate: "",
        //   materialCodes: "",
        //   remarks: "",
        // });
      } catch (error) {
        console.error("Error creating item:", error);
      } finally {
        setLoading(false);
      }
    })().catch((error) => {
      console.error("catch haha Error creating item:", error);
      navigate("/err/500", { replace: true });
    });
  };

  const handleApprove = (e: React.MouseEvent<HTMLButtonElement>): void => {
    // Update the row data in the list
    (async () => {
      try {
        setLoading(true);
        // Create a new list item :contentReference[oaicite:10]{index=10}
        /* const result =  */ await sp.web.lists
          .getById(listId)
          .items.getById(Number(formId))
          .update({ currStep: currStep + 1 });
        // console.log("Created item:", result);
        navigate("/");
      } catch (error) {
        console.error("Error creating item:", error);
      } finally {
        setLoading(false);
        // setRejectReason("");
        setHideDialog(true);
      }
    })().catch((error) => {
      console.error("catch haha Error creating item:", error);
      navigate("/err/500", { replace: true });
    });
  };

  const handleReject = (e: React.MouseEvent<HTMLButtonElement>): void => {
    // Update the row data in the list
    (async () => {
      try {
        setLoading(true);
        // Create a new list item :contentReference[oaicite:10]{index=10}
        /* const result =  */ await sp.web.lists
          .getById(listId)
          .items.getById(Number(formId))
          .update({
            redirectURL: "", // context.pageContext.web.absoluteUrl + context.pageContext.web.serverRelativeUrl + "#/initiater/",
            currStep: 0,
            reasonOfRejection: rejectReason,
            rejectedBy:
              context.pageContext.user.email ||
              context.pageContext.user.loginName ||
              context.pageContext.user.displayName ||
              "",
            rejectedFromStep: currStep,
          });
        // console.log("Created item:", result);
        navigate("/", { replace: true });
      } catch (error) {
        console.error("Error creating item:", error);
      } finally {
        setLoading(false);
        // setRejectReason("");
        setHideDialog(true);
      }
    })().catch((error) => {
      console.error("catch haha Error creating item:", error);
      navigate("/err/500", { replace: true });
    });
  };

  const options: IComboBoxOption[] = [
    { key: "umargam", text: "Umargam" },
    // { key: "tumb", text: "Tumb" },
    { key: "silvassa", text: "Silvassa" },
  ];
  const optionsPl: IComboBoxOption[] = [
    { key: "pl001", text: "PL001" },
    { key: "pl002", text: "PL002" },
    { key: "pl003", text: "PL003" },
  ];

  const statusMsg = (() => {
    switch (currStep) {
      case 1:
        return "GM User Approval Stage";
      case 2:
        return "PP Dept Approval Stage";
      case 3:
        return "QC Dept Approval Stage";
      case 0: {
        if (formId) {
          return "Update the Request";
        }
        return "Create new Request";
      }
      default:
        return "Create new Request";
    }
  })();

  const formEditMode = formId !== undefined;
  const newForm = currStep === 0;
  const disableSubmit =
    formDetails.location === "" ||
    formDetails.plantCode === "" ||
    formDetails.startDate === "" ||
    formDetails.remarks === "" ||
    loading;
  return (
    <div style={{ margin: "auto", maxWidth: "600px" }}>
      <Dialog
        hidden={hideDialog}
        onDismiss={toggleHideDialog}
        dialogContentProps={dialogContentProps}
      >
        <DialogContent>
          <TextField
            value={rejectReason}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) =>
              setRejectReason(e.target.value)
            }
            label="Reason of Rejection"
            multiline
            /* rows={12} */ name="reasonOfRejection"
          />
        </DialogContent>
        <DialogFooter>
          <PrimaryButton
            onClick={handleReject}
            disabled={!rejectReason}
            text="Reject Form"
          />
          <DefaultButton onClick={toggleHideDialog} text="Cancel" />
        </DialogFooter>
      </Dialog>
      {loading ? (
        <p>Working...</p>
      ) : (
        <form action="" onSubmit={handleSubmit}>
          <h4>
            Status: <u>{statusMsg}</u>
          </h4>
          <Stack horizontal tokens={stackTokens} horizontalAlign="stretch">
            <ComboBox
              // defaultSelectedKey="C"
              disabled={!newForm}
              selectedKey={comBoxSelectedKey.location}
              onChange={(e, opt) => {
                const name = "location";
                const value = opt?.text || "";
                setComBoxSelectedKey((p) => ({
                  ...p,
                  location: String(opt?.key) || "",
                }));
                setFormDetails((prevDetails) => ({
                  ...prevDetails,
                  [name]: value,
                }));
              }}
              label="Location"
              options={options}
              // styles={comboBoxStyles}
            />
            <ComboBox
              // defaultSelectedKey="C"
              disabled={!newForm}
              selectedKey={comBoxSelectedKey.plantCode}
              onChange={(e, opt) => {
                const name = "plantCode";
                const value = opt?.text || "";
                setComBoxSelectedKey((p) => ({
                  ...p,
                  plantCode: String(opt?.key) || "",
                }));
                setFormDetails((prevDetails) => ({
                  ...prevDetails,
                  [name]: value,
                }));
              }}
              label="Plant Code & Name"
              options={optionsPl}
              // styles={comboBoxStyles}
            />
            <TextField
              disabled={!newForm}
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
            disabled={!newForm}
            value={formDetails.remarks}
            onChange={handleChange}
            label="Remarks"
            multiline
            /* rows={12} */ name="remarks"
          />

          {!newForm ? (
            <>
              {currStep > 1 ? (
                <div>
                  <span>Approved By:</span>
                  <ol>
                    {currStep > 1 && <li>GM User</li>}
                    {currStep > 2 && <li>PP Department</li>}
                    {currStep > 3 && <li>QC Department</li>}
                  </ol>
                </div>
              ) : null}

              <Stack
                horizontal
                tokens={{ childrenGap: 8 }}
                style={{ marginTop: "12px", float: "right" }}
              >
                <DefaultButton
                  type="button"
                  text="Reject"
                  onClick={toggleHideDialog}
                />
                <PrimaryButton
                  type="button"
                  text="Approve"
                  onClick={handleApprove}
                />
              </Stack>
            </>
          ) : (
            <PrimaryButton
              text={
                loading
                  ? formEditMode
                    ? "Updating..."
                    : "Submitting..."
                  : formEditMode
                  ? "Update"
                  : "Submit"
              }
              disabled={disableSubmit}
              type="submit"
              style={{ marginTop: "12px", float: "right" }}
            />
          )}
        </form>
      )}
    </div>
  );
};

export default InitiaterForm;
