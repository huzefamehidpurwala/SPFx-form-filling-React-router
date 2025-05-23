import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";
import { Icon, PrimaryButton } from "@fluentui/react";
import { Link, useNavigate } from "react-router-dom";
import { sp } from "@pnp/sp";
import { listId } from "./Form";
import styles from "../WorkflowBillApproval.module.scss";
import { ContextStore } from "../Context/ContextStore";

export interface IDetailsListBasicExampleItem {
  key: number;
  isAuthorized: number;
  [key: string]: string | number;
}

// Columns definition (unchanged from class version)
const columns: IColumn[] = [
  {
    key: "column1",
    name: "Form ID",
    fieldName: "Id",
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
    onRender(item) {
      return (
        <>
          {Boolean(item.isAuthorized) ? (
            <div className={styles.links}>
              <Link
                to={"/form/" + item.Id}
                style={{ textDecoration: "underline" }}
              >
                {item.Id} <Icon iconName="NavigateExternalInline" />
              </Link>
            </div>
          ) : (
            <span>{item.Id}</span>
          )}
        </>
      );
    },
  },
  {
    key: "column2",
    name: "Location",
    fieldName: "location",
    maxWidth: 200,
    minWidth: 100,
    isResizable: true,
  },
  {
    key: "column3",
    name: "Plant Code",
    fieldName: "plantCode",
    maxWidth: 200,
    minWidth: 100,
    isResizable: true,
  },
  {
    key: "column4",
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
          <span
            style={{
              color: item.currStep === 0 ? "red" : undefined,
            }}
          >
            {statusMsg}
          </span>
        </div>
      );
    },
  },
];

const FormsList: React.FC = () => {
  const { spContext: context } = React.useContext(ContextStore);

  const navigate = useNavigate();

  // State hooks for items and selectionDetails
  const [items, setItems] = React.useState<IDetailsListBasicExampleItem[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    (async () => {
      try {
        setLoading(true);
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

        setItems(
          itemList.map((item) => {
            const rCurrStep = item.currStep as number;
            const autherEMail = (item as any).Author.EMail as string;

            let isAuthorized = false;
            switch (rCurrStep) {
              case 1:
                isAuthorized = userGroups.some(
                  (group) =>
                    group.Title.toLowerCase() ===
                    "gm " + item.location.toLowerCase()
                );
                break;

              case 2:
                isAuthorized = userGroups.some(
                  (group) => group.Title === "PP Department"
                );
                break;

              case 3:
                isAuthorized = userGroups.some(
                  (group) => group.Title === "GM QC"
                );
                break;

              case 0: {
                // isAuthorized = (context.pageContext.user.email.toLowerCase() || context.pageContext.user.loginName.toLowerCase()).includes(authEMail.toLowerCase());
                isAuthorized =
                  autherEMail.toLowerCase() ===
                  (context.pageContext.user.email.toLowerCase() ||
                    context.pageContext.user.loginName.toLowerCase());
                break;
              }

              default:
                break;
            }

            return {
              ...item,
              key: item.Id,
              isAuthorized: Number(isAuthorized),
            };
          })
        );
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
  }, []);

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
      {loading ? (
        <p>Fetching Details...</p>
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
            items={items}
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
          {/* Optionally display selectionDetails somewhere, e.g.: */}
          {/* <div>{selectedId}</div> */}
        </>
      )}
    </div>
  );
};

export default FormsList;
