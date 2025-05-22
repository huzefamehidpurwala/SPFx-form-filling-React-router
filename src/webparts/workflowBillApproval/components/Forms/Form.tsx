import * as React from "react";
import {
  ComboBox,
  DatePicker,
  DefaultButton,
  defaultDatePickerStrings,
  Dialog,
  DialogContent,
  DialogFooter,
  DialogType,
  IComboBoxOption,
  IStackTokens,
  mergeStyleSets,
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
export const listId = "86892207-d198-453b-9b56-01044bc52533"; // Form Entry

const dialogContentProps = {
  type: DialogType.largeHeader,
  title: "Confirm Rejection",
};

const ppComments = [
  "All Material Code are existing in requested Plant code",
  "Alternate Material is properly define in Bill of Material",
  "Storage Location properly define in Bill of Material",
  "Work Center in Routing check Validation and found OK",
  "Work Centre Cost Center check for Validation and found OK",
  "All Other Parameters in Bill of Material Checked",
];
const qcComments = [
  "Quality Check is / active & maintained for All the Material in Bill of Material",
  "QC Parameters are properly maintained in System Material master Tab for All Material in BOM",
  "All BOM Component Quantity is check and Found O.K",
];

const datePickerStyles = mergeStyleSets({
  root: { selectors: { "> *": { marginBottom: 15 } } },
  control: { maxWidth: 300, marginBottom: 15 },
});

const Form: React.FC = () => {
  const { spContext: context } = React.useContext(ContextStore);

  const { formId } = useParams();
  const navigate = useNavigate();

  const [formDetails, setFormDetails] = React.useState<IFormDetails>({
    location: "",
    plantCode: "",
    startDate: new Date().toISOString(),
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
  const [currStep, setCurrStep] = React.useState<number>(-2); // formId !== undefined ? 0 : -1
  const [usrGroups, setUsrGroups] = React.useState<{ Title: string }[]>([]);
  const [comments, setComments] = React.useState<Record<number, boolean>>({});
  const [authEMail, setAuthEMail] = React.useState(""); // purposely EMail has m capital as same is the property in response

  const toggleHideDialog = (): void => setHideDialog((p) => !p);

  React.useEffect(() => {
    (async () => {
      try {
        setLoading(true);
        await (async () => {
          // Get all groups for the current user
          const userGroups: Array<{ Title: string }> =
            await sp.web.currentUser.groups.select("Title").get();
          setUsrGroups(userGroups);
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

    if (!formId) {
      // navigate("/err/404");
      setCurrStep(-1);
      return;
    }

    // Fetch the form data
    (async () => {
      try {
        setLoading(true);
        await (async () => {
          // Create a new list item
          const result = await sp.web.lists
            .getById(listId)
            .items.getById(Number(formId))
            .select(
              "Author/EMail",
              "location",
              "plantCode",
              "startDate",
              "remarks",
              "currStep",
              "Id"
            )
            .expand("Author")
            .get();
          //   console.log("Created item:", result);
          setFormDetails(result);
          setCurrStep(result.currStep);
          setAuthEMail(result.Author.EMail);
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

  React.useEffect(() => {
    if (formId && usrGroups.length > 0 && currStep > -1) {
      let isAuthorized = false;
      switch (currStep) {
        case 1:
          isAuthorized = usrGroups.some((group) =>
            group.Title.toLowerCase().includes(
              formDetails.location.toLowerCase()
            )
          );
          break;

        case 2:
          isAuthorized = usrGroups.some((group) =>
            group.Title.toLowerCase().includes("pp")
          );
          break;

        case 3:
          isAuthorized = usrGroups.some((group) =>
            group.Title.toLowerCase().includes("qc")
          );
          break;

        case 0: {
          isAuthorized = (
            context.pageContext.user.email.toLowerCase() ||
            context.pageContext.user.loginName.toLowerCase()
          ).includes(authEMail.toLowerCase());
          break;
        }

        default:
          break;
      }

      if (!isAuthorized) {
        navigate("/err/401", { replace: true });
      }
    }
  }, [usrGroups.length, currStep, formId]);

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
        // Create a new list item
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
          navigate("/succ/sub", { replace: true });
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
        navigate("/succ/upd", { replace: true });
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
      const isCommentsEmpty = Object.keys(comments).every(
        (key) => !comments[Number(key)]
      );
      const sendComments = !isCommentsEmpty
        ? ppComments.filter((_, i) => comments[i]).join("\n")
        : "Approved by " + context.pageContext.user.displayName;
      try {
        setLoading(true);
        // Create a new list item
        /* const result =  */ await sp.web.lists
          .getById(listId)
          .items.getById(Number(formId))
          .update({
            currStep: currStep + 1,
            comments: sendComments,
          });
        // console.log("Created item:", result);
        navigate("/succ/app", { replace: true });
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
        // Create a new list item
        /* const result =  */ await sp.web.lists
          .getById(listId)
          .items.getById(Number(formId))
          .update({
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
        setRejectReason("");
        navigate("/succ/rej", { replace: true });
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

  const onFormatDate = (date?: Date): string => {
    return !date
      ? ""
      : String(date.getDate()).padStart(2, "0") +
          "/" +
          String(date.getMonth() + 1).padStart(2, "0") +
          "/" +
          date.getFullYear();
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

  const extras = (() => {
    switch (currStep) {
      case 2:
        return (
          <div
            style={{
              marginTop: "16px",
              marginBottom: "16px",
              paddingBottom: "16px",
              borderBottom: "1px solid black",
            }}
          >
            {ppComments.map((st, i) => {
              return (
                <div key={i}>
                  <input
                    type="checkbox"
                    checked={comments[i]}
                    id={i.toString()}
                    name={i.toString()}
                    onChange={(e) =>
                      setComments((p) => ({ ...p, [i]: e.target.checked }))
                    }
                  />
                  <label htmlFor={i.toString()}>{st}</label>
                </div>
              );
            })}
          </div>
        );

      case 3:
        return (
          <div
            style={{
              marginTop: "16px",
              paddingBottom: "16px",
              borderBottom: "1px solid black",
            }}
          >
            {qcComments.map((st, i) => {
              return (
                <div key={i}>
                  <input
                    type="checkbox"
                    checked={comments[i]}
                    id={i.toString()}
                    name={i.toString()}
                    onChange={(e) =>
                      setComments((p) => ({ ...p, [i]: e.target.checked }))
                    }
                  />
                  <label htmlFor={i.toString()}>{st}</label>
                </div>
              );
            })}
          </div>
        );

      case 1:
      case -1:
      case 0:
      default:
        return null;
    }
  })();
  const statusMsg = (() => {
    switch (currStep) {
      case 1:
        return "GM User Approval Stage";
      case 2:
        return "PP Dept Approval Stage";
      case 3:
        return "QC Dept Approval Stage";
      case -1:
        return "Create New Request";
      case 0:
        return "Update the Request";
      default:
        return "Approved";
    }
  })();

  const updateMode = currStep === 0;
  const isFormDisabled = currStep > 0;
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
        modalProps={{ isBlocking: true }}
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
              disabled={isFormDisabled}
              selectedKey={comBoxSelectedKey.location}
              onChange={(_, opt) => {
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
            />
            <ComboBox
              disabled={isFormDisabled}
              selectedKey={comBoxSelectedKey.plantCode}
              onChange={(_, opt) => {
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
            />
            <DatePicker
              disabled={isFormDisabled}
              label="Start date"
              ariaLabel="Select a date. Input format is dd/mm/yyyy."
              allowTextInput // used just for good UI to render as input element
              value={new Date(formDetails.startDate)}
              onSelectDate={(date) => {
                if (date) {
                  setFormDetails((p) => ({
                    ...p,
                    startDate: date.toISOString(),
                  }));
                }
              }}
              formatDate={onFormatDate}
              // minDate={new Date()}
              className={datePickerStyles.control}
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
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
            disabled={isFormDisabled}
            value={formDetails.remarks}
            onChange={handleChange}
            label="Remarks"
            multiline
            /* rows={12} */ name="remarks"
          />

          {isFormDisabled ? (
            <>
              {extras}

              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "flex-start",
                  marginTop: "16px",
                }}
              >
                <div>
                  {currStep > 1 ? (
                    <>
                      <span>Approved By:</span>
                      <ol>
                        {currStep > 1 && <li>GM User</li>}
                        {currStep > 2 && <li>PP Department</li>}
                        {currStep > 3 && <li>QC Department</li>}
                      </ol>
                    </>
                  ) : null}
                </div>

                <Stack
                  horizontal
                  tokens={{ childrenGap: 8 }}
                  // style={{ marginTop: "12px", float: "right" }}
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
              </div>
            </>
          ) : (
            <PrimaryButton
              text={
                loading
                  ? updateMode
                    ? "Updating..."
                    : "Submitting..."
                  : updateMode
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

export default Form;
