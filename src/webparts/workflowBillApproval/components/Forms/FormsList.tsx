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

export interface IDetailsListBasicExampleItem {
  key: number;
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
        <div className={styles.links}>
          <Link to={"/form/" + item.Id} style={{ textDecoration: "underline" }}>
            {item.Id} <Icon iconName="NavigateExternalInline" />
          </Link>
        </div>
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
          <span
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
          />
          <span>{statusMsg}</span>
        </div>
      );
    },
  },
];

const FormsList: React.FC = () => {
  const navigate = useNavigate();

  // State hooks for items and selectionDetails
  const [items, setItems] = React.useState<IDetailsListBasicExampleItem[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    (async () => {
      try {
        setLoading(true);
        await (async () => {
          // Create a new list item
          const result = (await sp.web.lists
            .getById(listId)
            .items.select("Id", "location", "plantCode", "currStep")
            .filter("currStep le 3")
            .get()) as {
            Id: number;
            location: string;
            plantCode: string;
            currStep: number;
          }[];
          setItems(result.map((item) => ({ key: item.Id, ...item })));
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
