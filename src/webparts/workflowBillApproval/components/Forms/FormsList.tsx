import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";
import { FontSizes, Icon, PrimaryButton, SpinnerSize } from "@fluentui/react";
import { Link, useNavigate } from "react-router-dom";
import { sp } from "@pnp/sp";
import { listId } from "./Form";
import { ContextStore } from "../Context/ContextStore";
import styles from "../WorkflowBillApproval.module.scss";
import { useQuery } from "@tanstack/react-query";
import Loading from "../../../helper/Loading";

export interface IDetailsListBasicExampleItem {
  key: number;
  isAuthorized: number;
  [key: string]: string | number;
}

// Columns definition (unchanged from class version)
const columns: IColumn[] = [
  {
    key: "column1",
    name: "Edit Button",
    fieldName: "editBtn",
    minWidth: 16,
    maxWidth: 16,
    isIconOnly: true,
    isResizable: false,
    onRender(item) {
      return (
        <>
          {Boolean(item.isAuthorized) ? (
            <div className={styles.links}>
              <Link to={"/form/" + item.Id}>
                <Icon
                  iconName="NavigateExternalInline" // iconProps={{ iconName: "NavigateExternalInline" }}
                  style={{
                    fontSize: FontSizes.size16, // e.g. “24px” under the hood
                    width: FontSizes.size16,
                    height: FontSizes.size16,
                    textDecoration: "none",
                  }}
                />
              </Link>
            </div>
          ) : null}
        </>
      );
    },
  },
  {
    key: "column2",
    name: "Form ID",
    fieldName: "Id",
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
  },
  {
    key: "column3",
    name: "Location",
    fieldName: "location",
    maxWidth: 200,
    minWidth: 100,
    isResizable: true,
  },
  {
    key: "column4",
    name: "Plant Code",
    fieldName: "plantCode",
    maxWidth: 200,
    minWidth: 100,
    isResizable: true,
  },
  {
    key: "column5",
    name: "Status",
    fieldName: "status",
    maxWidth: 300,
    minWidth: 100,
    isResizable: true,
    onRender(item) {
      let statusMsg = "";
      switch (item.currStep) {
        case 1:
          statusMsg = "GM User Approval Stage";
          break;

        case 2:
          statusMsg = "PP Dept Approval Stage";
          break;

        case 3:
          statusMsg = "QC Dept Approval Stage";
          break;

        case -1:
          statusMsg = "Create New Request";
          break;

        case 0:
          statusMsg = "Rejected";
          break;

        default:
          statusMsg = "Approved";
          break;
      }
      return (
        <div
          style={{
            display: "flex",
            columnGap: "8px",
            // justifyContent: "center",
            alignItems: "center",
          }}
        >
          {/* <span
            style={{
              width: "12px",
              height: "12px",
              backgroundColor:
                item.currStep === 0
                  ? "red"
                  : item.currStep > 3
                  ? "green"
                  : undefined,
              borderRadius: "999px",
            }}
          /> */}
          <span style={{ color: item.currStep === 0 ? "red" : undefined }}>
            {statusMsg}
          </span>
        </div>
      );
    },
  },
];

const fetchFormList = async (
  currUserEmail: string
): Promise<IDetailsListBasicExampleItem[]> => {
  const itemListPr = sp.web.lists
    .getById(listId)
    .items.expand("Author")
    .select("Author/EMail", "Id", "location", "plantCode", "currStep")
    .filter("currStep le 3")
    .get();

  // Get all groups for the current user
  const userGroupsPr = sp.web.currentUser.groups.select("Title").get(); // : Array<{ Title: string }>

  const [itemList, userGroups] = (await Promise.all([
    itemListPr,
    userGroupsPr,
  ])) as [
    {
      Id: number;
      location: string;
      plantCode: string;
      currStep: number;
    }[],
    { Title: string }[]
  ];

  // as one user can be in multiple groups
  const currUserGrpNum: number[] = [];
  let currUserGmLocation = "";
  userGroups.forEach((grp) => {
    if (grp.Title === "GM QC") {
      currUserGrpNum.push(3);
    } else if (grp.Title === "PP Department") {
      currUserGrpNum.push(2);
    } else if (grp.Title.toLowerCase().includes("gm")) {
      currUserGrpNum.push(1);
      currUserGmLocation = grp.Title.toLowerCase().replace("gm ", "").trim();
    }
  });

  return itemList.map((item) => {
    const { Id, location, currStep } = item;
    const authorEmail = (item as any).Author?.EMail?.toLowerCase() || "";
    const step = currStep as number;

    const isGroupMatch = currUserGrpNum.includes(step);
    const isGmLocationMatch =
      step === 1 &&
      isGroupMatch &&
      currUserGmLocation === location.toLowerCase();

    const isAuthorized =
      step === 0
        ? authorEmail === currUserEmail.toLowerCase()
        : step === 1
        ? isGmLocationMatch
        : isGroupMatch;

    return {
      ...item,
      key: Id,
      isAuthorized: Number(isAuthorized),
    };
  });
};

const FormsList: React.FC = () => {
  const { spContext: context } = React.useContext(ContextStore);

  const navigate = useNavigate();

  const {
    isInitialLoading: loading,
    data: items,
    isError,
    error,
    isFetching,
  } = useQuery(["formLists"], () =>
    fetchFormList(
      context.pageContext.user.email || context.pageContext.user.loginName
    )
  );

  // Error handling
  if (isError && error instanceof Error) {
    console.error("Error fetching form list:", error);
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        <h2>Error fetching forms</h2>
        <p>{(error as Error).message}</p>
        <PrimaryButton
          onClick={() => window.location.reload()}
          iconProps={{ iconName: "Refresh" }}
        >
          Refresh
        </PrimaryButton>
      </div>
    );
  }

  // Render output
  return (
    <div
      style={{
        maxWidth: "900px",
        margin: "0 auto",
        display: "flex",
        flexDirection: "column",
        gap: "20px",
      }}
    >
      {loading && !items?.length ? (
        <Loading label="Fetching..." />
      ) : (
        <>
          <div style={{ marginLeft: "auto" }}>
            <PrimaryButton
              onClick={() => navigate("/form")}
              iconProps={{ iconName: "Add" }}
            >
              New Request
            </PrimaryButton>
          </div>
          <DetailsList
            items={items!}
            columns={columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            // selection={selection}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="select row"
            selectionMode={SelectionMode.none}
          />
          {!loading && isFetching && (
            <Loading label="Updating..." size={SpinnerSize.small} />
          )}
        </>
      )}
    </div>
  );
};

export default FormsList;
