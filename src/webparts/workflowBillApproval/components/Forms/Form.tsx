import * as React from "react";
import {
  Checkbox,
  ComboBox,
  DatePicker,
  DefaultButton,
  defaultDatePickerStrings,
  Dialog,
  DialogContent,
  DialogFooter,
  DialogType,
  Dropdown,
  IComboBoxOption,
  IDropdownOption,
  IStackTokens,
  mergeStyleSets,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";
import { sp } from "@pnp/sp";
import { ContextStore } from "../Context/ContextStore";
import { useNavigate, useParams } from "react-router-dom";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Loading from "../../../helper/Loading";

export type IFormDetails = {
  location: string;
  plantCode: string;
  startDate: string;
  materialCodes: string;
  billParameters: Record<number, boolean>;
  remarks: string;
  bomRequest: string;
};

export const stackTokens: IStackTokens = { childrenGap: 40 };
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
export const comments = [
  "All Material Codes are Existing in requested Plant code",
  "For New Material BOM Routing Attached",
  "Bill Of Material & Routing Attached",
];

const locationOptions: IComboBoxOption[] = [
  { key: "Umargam", text: "Umargam" },
  { key: "Silvassa", text: "Silvassa" },
];
const plantOptions: IComboBoxOption[] = [
  { key: "PL001", text: "PL001" },
  { key: "PL002", text: "PL002" },
  { key: "PL003", text: "PL003" },
];
const dropdownOptions: IDropdownOption[] = [
  {
    key: "New Material Bill Of Material",
    text: "New Material Bill Of Material",
  },
  {
    key: "Changes in Existing Bill of Material",
    text: "Changes in Existing Bill of Material",
  },
];

const datePickerStyles = mergeStyleSets({
  root: { selectors: { "> *": { marginBottom: 15 } } },
  control: { maxWidth: 300, marginBottom: 15 },
});

/**
 * fetchFormItem
 *
 * This function calls PnP/SP to get the list item and the current user's groups,
 * then determines authorization and returns the list item and its current step.
 *
 * @param formId - the numeric ID of the list item to fetch
 * @param context - the SP context (to get current user and web)
 */
async function fetchFormItem(
  formId: number,
  context: WebPartContext
): Promise<{ listItem: any; currStep: number }> {
  // 1) Query the list item by its ID, expanding the "Author" field so we can read the creator’s email.
  const listItemPromise = sp.web.lists
    .getById(listId)
    .items.getById(formId)
    .expand("Author", "Editor")
    .select(
      "Author/EMail",
      "Editor/Title",
      "location",
      "plantCode",
      "startDate",
      "remarks",
      "currStep",
      "materialCodes",
      "billParameters",
      "bomRequest",
      "Id"
    )
    .get();

  // 2) Query the groups for the current user
  const userGroupsPromise = sp.web.currentUser.groups.select("Title").get();

  // Wait for both to resolve in parallel
  const [listItemResult, userGroups] = await Promise.all([
    listItemPromise,
    userGroupsPromise,
  ]);

  const rCurrStep = listItemResult.currStep as number;
  const authorEmail = listItemResult.Author.EMail as string;

  // 3) Check if the current user is authorized for this step
  let isAuthorized = false;
  switch (rCurrStep) {
    case 1:
      isAuthorized = userGroups.some(
        (grp) =>
          grp.Title.toLowerCase() ===
          "gm " + listItemResult.location.toLowerCase()
      );
      break;
    case 2:
      isAuthorized = userGroups.some((grp) => grp.Title === "PP Department");
      break;
    case 3:
      isAuthorized = userGroups.some((grp) => grp.Title === "GM QC");
      break;
    case 0: {
      // If currStep is 0 (meaning “rejected from a prior step”), only the original author can edit
      const currentUserEmail =
        context.pageContext.user.email || context.pageContext.user.loginName;
      isAuthorized =
        authorEmail.toLowerCase() === currentUserEmail.toLowerCase();
      break;
    }
    default:
      break;
  }

  if (!isAuthorized) {
    // If the user is not authorized, throw an error that our component can catch.
    // We’ll navigate to a 401 page in the component based on this.
    throw new Error("Unauthorized");
  }

  // Return both the raw list item and the current step number
  return { listItem: listItemResult, currStep: rCurrStep };
}

/**
 * useFormItemQuery
 *
 * Custom hook wrapping React Query's useQuery to fetch a form item by ID.
 *
 * @param formId - numeric ID of the item
 * @param context - SP context from React Context
 */
function useFormItemQuery(formId: number | undefined, context: WebPartContext) {
  return useQuery(
    ["formItem", formId],
    // queryFn only runs if formId is defined
    () => fetchFormItem(formId as number, context),
    {
      enabled: formId !== undefined, // only run when formId exists
      retry: false, // don’t retry on Unauthorized; let us handle it
    }
  );
}

/**
 * useCreateOrUpdateFormMutation
 *
 * Custom hook wrapping React Query's useMutation to handle both “create new”
 * and “update existing” scenarios in a single function.
 *
 * It inspects whether formId is present; if not, it does an “add” call; otherwise, it does an “update” call.
 *
 * @param formId - numeric ID (or undefined if new)
 * @param context - SP context to get pageContext for URLs
 * @param onSuccess - callback after success (e.g. navigate)
 */
function useCreateOrUpdateFormMutation(
  formId: number | undefined,
  context: WebPartContext,
  onSuccess: () => void
) {
  const queryClient = useQueryClient();

  return useMutation(
    async (formData: IFormDetails) => {
      // Build the common payload
      const absoluteUrl = context.pageContext.web.absoluteUrl;
      const pageRelativePath =
        context.pageContext.site.serverRequestPath.replace(
          context.pageContext.site.serverRelativeUrl,
          ""
        );
      const hashRoute = "#/form/";

      const payload = {
        location: formData.location,
        plantCode: formData.plantCode,
        startDate: formData.startDate,
        remarks: formData.remarks,
        billParameters: JSON.stringify(formData.billParameters),
        materialCodes: formData.materialCodes,
        bomRequest: formData.bomRequest,
        redirectURL: absoluteUrl + pageRelativePath + hashRoute,
        currStep: 1,
        reasonOfRejection: null,
        rejectedBy: null,
        rejectedFromStep: null,
      };

      if (!formId) {
        // If there’s no formId, we “create” a new item
        return sp.web.lists.getById(listId).items.add(payload);
      } else {
        // Otherwise, we “update” the existing item
        return sp.web.lists
          .getById(listId)
          .items.getById(formId)
          .update(payload);
      }
    },
    {
      onSuccess: () => {
        // Invalidate any “formItem” queries so stale data is not shown
        queryClient
          .invalidateQueries(["formItem", formId])
          .catch(console.error);
        onSuccess(); // e.g. navigate to success page
      },
    }
  );
}

/**
 * useApproveFormMutation
 *
 * Custom hook wrapping React Query's useMutation to “approve” the form:
 * advances currStep by +1 and writes any comments selected.
 *
 * @param formId - ID of the item
 * @param currStep - current step number (so we can +1)
 * @param comments - aggregated comments string
 * @param onSuccess - callback after successful approval
 */
function useApproveFormMutation(
  formId: number,
  currStep: number,
  comments: string,
  onSuccess: () => void
) {
  const queryClient = useQueryClient();

  return useMutation(
    async () => {
      return sp.web.lists
        .getById(listId)
        .items.getById(formId)
        .update({
          currStep: currStep + 1,
          comments: comments,
        });
    },
    {
      onSuccess: () => {
        queryClient
          .invalidateQueries(["formItem", formId])
          .catch(console.error);
        onSuccess();
      },
    }
  );
}

/**
 * useRejectFormMutation
 *
 * Custom hook wrapping React Query's useMutation to “reject” the form:
 * sets currStep to 0 and writes rejection info.
 *
 * @param formId - ID of item
 * @param currStep - the step from which we are rejecting
 * @param rejectReason - the text reason
 * @param rejectedBy - e.g. user email
 * @param onSuccess - callback after successful rejection
 */
function useRejectFormMutation(
  formId: number,
  currStep: number,
  rejectReason: string,
  rejectedBy: string,
  onSuccess: () => void
) {
  const queryClient = useQueryClient();

  return useMutation(
    async () => {
      return sp.web.lists.getById(listId).items.getById(formId).update({
        currStep: 0,
        reasonOfRejection: rejectReason,
        rejectedBy: rejectedBy,
        rejectedFromStep: currStep,
      });
    },
    {
      onSuccess: () => {
        queryClient
          .invalidateQueries(["formItem", formId])
          .catch(console.error);
        onSuccess();
      },
    }
  );
}

/**
 * useDeleteFormMutation
 *
 * Custom hook wrapping React Query's useMutation to delete the form item.
 *
 * @param formId - ID of item
 * @param onSuccess - callback after delete
 */
function useDeleteFormMutation(formId: number, onSuccess: () => void) {
  const queryClient = useQueryClient();

  return useMutation(
    async () => {
      return sp.web.lists.getById(listId).items.getById(formId).delete();
    },
    {
      onSuccess: () => {
        queryClient
          .invalidateQueries(["formItem", formId])
          .catch(console.error);
        onSuccess();
      },
    }
  );
}

/**
 * Form Component
 *
 * - Uses React Query's useQuery to fetch existing data (if editing).
 * - Uses multiple useMutation hooks to handle create, update, approve, reject, delete.
 * - Disables form fields or shows approval UI based on currStep.
 * - Comments explain why each hook/pattern is used.
 */
const Form: React.FC = () => {
  const { spContext: context } = React.useContext(ContextStore);
  const { formId: rawFormId } = useParams<{ formId: string }>();
  const navigate = useNavigate();
  // const queryClient = useQueryClient();

  // Parse formId from URL params; if missing or not a number, formIdNumber is undefined
  const formIdNumber = rawFormId ? Number(rawFormId) : undefined;

  // Local React state for form fields (controlled inputs)
  const [formDetails, setFormDetails] = React.useState<IFormDetails>({
    location: "",
    plantCode: "",
    startDate: "",
    materialCodes: "",
    billParameters: [],
    remarks: "",
    bomRequest: "",
  });
  const [commentsChecked, setCommentsChecked] = React.useState<
    Record<number, boolean>
  >({});
  const [rejectReason, setRejectReason] = React.useState("");
  const [itemEditedByName, setItemEditedByName] = React.useState("");

  // Dialog state for rejection modal
  const [hideDialog, setHideDialog] = React.useState(true);

  // 1) Fetch existing form data if formIdNumber is defined
  //    useFormItemQuery will internally call fetchFormItem, check authorization, etc.
  const {
    data: fetchedData,
    isInitialLoading: isLoading,
    isError: fetchError,
    error: fetchErrorObject,
  } = useFormItemQuery(formIdNumber, context);

  // current step (rCurrStep) will come from fetchedData if editing; otherwise default to -1 for “new”
  const currStep = fetchedData ? fetchedData.currStep : -1;

  // Populate formDetails state when fetchedData becomes available
  React.useEffect(() => {
    if (fetchedData) {
      const { listItem } = fetchedData;
      setFormDetails({
        location: listItem.location || "",
        plantCode: listItem.plantCode || "",
        startDate: listItem.startDate || "",
        materialCodes: listItem.materialCodes || "",
        remarks: listItem.remarks || "",
        bomRequest: listItem.bomRequest || "",
        billParameters: JSON.parse(listItem.billParameters || "{}"),
      });
      setItemEditedByName(listItem.Editor?.Title || "");
    }
  }, [fetchedData]);

  // If fetchError is “Unauthorized”, send to 401; otherwise, to 500
  React.useEffect(() => {
    if (fetchError && fetchErrorObject instanceof Error) {
      if (fetchErrorObject.message === "Unauthorized") {
        navigate("/err/401", { replace: true });
      } else {
        navigate("/err/500", { replace: true });
      }
    }
  }, [fetchError, fetchErrorObject, navigate]);

  // 2) Create/Update Mutation
  //    onSuccess we navigate to appropriate “success” URL
  const createOrUpdateMutation = useCreateOrUpdateFormMutation(
    formIdNumber,
    context,
    () => {
      if (formIdNumber) {
        navigate("/succ/upd", { replace: true });
      } else {
        navigate("/succ/sub", { replace: true });
      }
    }
  );

  // 3) Approve Mutation (only valid if editing and currStep > 0)
  const approveMutation = useApproveFormMutation(
    formIdNumber as number,
    currStep,
    // aggregate comments (if none checked, default to “Approved by User”)
    Object.keys(commentsChecked).some((i) => commentsChecked[Number(i)])
      ? ppComments.filter((_, i) => commentsChecked[i]).join("\n")
      : "Approved by " + context.pageContext.user.displayName,
    () => {
      navigate("/succ/app", { replace: true });
    }
  );

  // 4) Reject Mutation
  const rejectMutation = useRejectFormMutation(
    formIdNumber as number,
    currStep,
    rejectReason,
    context.pageContext.user.email ||
      context.pageContext.user.loginName ||
      context.pageContext.user.displayName ||
      "",
    () => {
      setRejectReason("");
      navigate("/succ/rej", { replace: true });
    }
  );

  // 5) Delete Mutation (only when editing & in update mode)
  const deleteMutation = useDeleteFormMutation(formIdNumber as number, () => {
    navigate("/succ/del", { replace: true });
  });

  // Helper to toggle reject dialog
  const toggleHideDialog = (): void => setHideDialog((old) => !old);

  // Format a Date object to “dd/MM/yyyy”
  const onFormatDate = (date?: Date): string => {
    if (!date) return "";
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };

  // Track whether we’re in “update” mode (currStep=0 means someone rejected and now they can edit)
  const updateMode = currStep === 0;
  // If currStep > 0 (approved stages), disable the form fields
  const isFormDisabled = currStep > 0;

  // Determine whether the submit/update button should be disabled
  const disableSubmit =
    !formDetails.location ||
    !formDetails.plantCode ||
    !formDetails.startDate ||
    !formDetails.remarks ||
    !formDetails.bomRequest ||
    !formDetails.materialCodes ||
    createOrUpdateMutation.isLoading;

  // If fetching data or running any mutation, show a “Working…” message
  if (
    (isLoading || createOrUpdateMutation.isLoading) &&
    !!formIdNumber &&
    !fetchedData
  ) {
    return <Loading label="Fetching..." />;
  }

  // Now render the form (either “new” or “edit”)
  return (
    <div style={{ margin: "auto", maxWidth: "600px" }}>
      {/* Rejection Dialog */}
      <Dialog
        hidden={hideDialog}
        onDismiss={toggleHideDialog}
        dialogContentProps={dialogContentProps}
        // modalProps={{ isBlocking: true }}
      >
        <DialogContent>
          <TextField
            value={rejectReason}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) =>
              setRejectReason(e.target.value)
            }
            label="Reason of Rejection"
            multiline
            name="reasonOfRejection"
          />
        </DialogContent>
        <DialogFooter>
          <PrimaryButton
            onClick={() => rejectMutation.mutate()}
            disabled={!rejectReason || rejectMutation.isLoading}
            text={rejectMutation.isLoading ? "Rejecting…" : "Reject Form"}
          />
          <DefaultButton onClick={toggleHideDialog} text="Cancel" />
        </DialogFooter>
      </Dialog>

      {/* Main Form */}
      <form
        onSubmit={(e) => {
          e.preventDefault();
          // Trigger create/update mutation
          createOrUpdateMutation.mutate({ ...formDetails });
        }}
      >
        {/* Status message based on currStep */}
        <h4>
          Status:{" "}
          <u>
            {(() => {
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
            })()}
          </u>
        </h4>

        <Stack horizontal tokens={stackTokens} horizontalAlign="stretch">
          {/* Location Dropdown */}
          <ComboBox
            disabled={isFormDisabled}
            selectedKey={formDetails.location}
            onChange={(_, opt) => {
              const value = opt?.text || "";
              setFormDetails((prev) => ({ ...prev, location: value }));
            }}
            label="Location"
            options={locationOptions}
          />

          {/* Plant Code Dropdown */}
          <ComboBox
            disabled={isFormDisabled}
            selectedKey={formDetails.plantCode}
            onChange={(_, opt) => {
              const value = opt?.text || "";
              setFormDetails((prev) => ({ ...prev, plantCode: value }));
            }}
            label="Plant Code & Name"
            options={plantOptions}
          />

          {/* Start Date Picker */}
          <DatePicker
            disabled={isFormDisabled}
            label="Start date"
            ariaLabel="Select a date. Input format is dd/mm/yyyy."
            allowTextInput // makes it render as an input for better UX
            value={
              formDetails.startDate
                ? new Date(formDetails.startDate)
                : undefined
            }
            onSelectDate={(date) => {
              if (date) {
                setFormDetails((p) => ({
                  ...p,
                  startDate: date.toISOString(),
                }));
              }
            }}
            formatDate={onFormatDate}
            className={datePickerStyles.control}
            strings={defaultDatePickerStrings}
          />
        </Stack>

        <div style={{ display: "flex", gap: 40 }}>
          {/* Dropdown Section */}
          <div>
            <Dropdown
              options={dropdownOptions}
              label={"Bill of Material Request for Uploading"}
              selectedKey={formDetails.bomRequest}
              disabled={isFormDisabled}
              onChange={(_, option) =>
                setFormDetails((p) => ({
                  ...p,
                  bomRequest: option?.key as string,
                }))
              }
              style={{ width: 250, marginTop: 8 }}
            />
          </div>
          {/* Checkbox Section */}
          <div>
            <label
              style={{
                fontWeight: "600",
                color: isFormDisabled ? "rgb(161, 159, 157)" : undefined,
              }}
            >
              Confirmation on Bill Of Material Parameters
            </label>
            <div style={{ marginTop: 8 }}>
              {comments.map((comm, i) => {
                return (
                  <React.Fragment key={i}>
                    <Checkbox
                      disabled={isFormDisabled}
                      label={comm}
                      checked={!!formDetails.billParameters[i]}
                      styles={{ root: { marginBottom: 6 } }}
                      onChange={(_, chckd) =>
                        setFormDetails((p) => ({
                          ...p,
                          billParameters: { ...p.billParameters, [i]: !!chckd },
                        }))
                      }
                    />
                  </React.Fragment>
                );
              })}
            </div>
          </div>
        </div>

        {/* Material Codes TextField */}
        <TextField
          disabled={isFormDisabled}
          multiline
          rows={6}
          value={formDetails.materialCodes}
          onChange={(e: any) => {
            const { name, value } = e.target;
            setFormDetails((prev) => ({ ...prev, [name]: value }));
          }}
          label="Material Codes ( Max. 12 nos. )"
          name="materialCodes"
        />

        {/* Remarks TextField */}
        <div
          style={
            currStep > 1
              ? {
                  paddingBottom: "16px",
                  borderBottom: "1px solid black",
                  marginBottom: "16px",
                }
              : undefined
          }
        >
          <TextField
            disabled={isFormDisabled}
            value={formDetails.remarks}
            onChange={(e: any) => {
              const { name, value } = e.target;
              setFormDetails((prev) => ({ ...prev, [name]: value }));
            }}
            label="Remarks"
            multiline
            name="remarks"
          />
        </div>

        {isFormDisabled ? (
          <>
            {/* Show checkboxes for PP / QC comments if we’re in those stages */}
            {currStep === 2 || currStep === 3 ? (
              <div>
                {(currStep === 2 ? ppComments : qcComments).map((text, idx) => (
                  <React.Fragment key={idx}>
                    <Checkbox
                      label={text}
                      styles={{ root: { marginBottom: 6 } }}
                      checked={!!commentsChecked[idx]}
                      onChange={(_, chckd) =>
                        setCommentsChecked((old) => ({
                          ...old,
                          [idx]: !!chckd,
                        }))
                      }
                    />
                  </React.Fragment>
                ))}
              </div>
            ) : null}

            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "flex-start",
                marginTop: "16px",
              }}
            >
              {/* Approved By List */}
              <div>
                {currStep > 1 ? (
                  <>
                    <span>Approved By:</span>
                    <ol>
                      {currStep > 1 && (
                        <li>
                          GM User
                          {currStep === 2 ? (
                            <>
                              : <u>{itemEditedByName}</u>
                            </>
                          ) : null}
                        </li>
                      )}
                      {currStep > 2 && (
                        <li>
                          PP Department
                          {currStep === 3 ? (
                            <>
                              : <u>{itemEditedByName}</u>
                            </>
                          ) : null}
                        </li>
                      )}
                      {currStep > 3 && <li>QC Department</li>}
                    </ol>
                  </>
                ) : null}
              </div>

              {/* Approve / Reject Buttons */}
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <DefaultButton
                  type="button"
                  text="Reject"
                  onClick={toggleHideDialog}
                  disabled={
                    approveMutation.isLoading || rejectMutation.isLoading
                  }
                />
                <PrimaryButton
                  type="button"
                  text={approveMutation.isLoading ? "Approving…" : "Approve"}
                  onClick={() => approveMutation.mutate()}
                  disabled={
                    approveMutation.isLoading || rejectMutation.isLoading
                  }
                />
              </Stack>
            </div>
          </>
        ) : (
          <>
            {/* Submit / Update Buttons */}
            <PrimaryButton
              text={
                createOrUpdateMutation.isLoading
                  ? updateMode
                    ? "Updating…"
                    : "Submitting…"
                  : updateMode
                  ? "Update"
                  : "Submit"
              }
              disabled={disableSubmit}
              type="submit"
              style={{ marginTop: "12px", float: "right" }}
            />
            {/* If in update mode (currStep=0), also show Delete button */}
            {updateMode && (
              <DefaultButton
                text={deleteMutation.isLoading ? "Deleting…" : "Delete"}
                disabled={deleteMutation.isLoading}
                type="button"
                onClick={() => deleteMutation.mutate()}
                style={{
                  marginTop: "12px",
                  marginRight: "8px",
                  float: "right",
                }}
              />
            )}
          </>
        )}
      </form>
    </div>
  );
};

export default Form;
