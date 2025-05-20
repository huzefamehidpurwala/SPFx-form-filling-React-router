import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";
import { MarqueeSelection } from "@fluentui/react/lib/MarqueeSelection";
import { PrimaryButton } from "@fluentui/react";
import { useNavigate } from "react-router-dom";
import { sp } from "@pnp/sp";
import { listId } from "./Form";

export interface IDetailsListBasicExampleItem {
  key: number;
  [key: string]: string | number;
}

// 3. Columns definition (unchanged from class version)
const columns: IColumn[] = [
  {
    key: "column1",
    name: "Form ID",
    fieldName: "Id",
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
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
  },
];

const FormsList: React.FC = () => {
  const navigate = useNavigate();

  // State hooks for items and selectionDetails
  const [items, setItems] = React.useState<IDetailsListBasicExampleItem[]>([]);
  const [loading, setLoading] = React.useState(true);

  // Memoized Selection instance so it's only created once
  const selection = React.useMemo(
    () =>
      new Selection({
        onSelectionChanged: () => {
          const selectedFormId = getSelectionDetails();
          if (selectedFormId) {
            navigate("/form/" + selectedFormId);
          }
        },
      }),
    []
  );

  // Helper to compute selection summary from the Selection instance
  function getSelectionDetails(): string | undefined {
    const selectionCount = selection.getSelectedCount();
    switch (selectionCount) {
      case 1: {
        const [selectedItem] =
          selection.getSelection() as IDetailsListBasicExampleItem[];
        return selectedItem.Id as string;
      }
      case 0:
      default:
        return undefined;
    }
  }

  React.useEffect(() => {
    (async () => {
      try {
        setLoading(true);
        await (async () => {
          // Create a new list item :contentReference[oaicite:10]{index=10}
          const result = (await sp.web.lists
            .getById(listId)
            .items.select("Id", "location", "plantCode", "currStep")
            .get()) as {
            Id: number;
            location: string;
            plantCode: string;
            currStep: number;
          }[];
          //   console.log("Created item:", result);
          setItems(
            result.map((item) => ({
              key: item.Id,
              ...item,
              status: ((currStep: number) => {
                switch (currStep) {
                  case 1:
                    return "GM User Approval Stage";
                  case 2:
                    return "PP Dept Approval Stage";
                  case 3:
                    return "QC Dept Approval Stage";
                  case -1:
                    return "Create New Request";
                  case 0: {
                    // if (formId) {
                    return "Update the Request";
                    // }
                    //   return "Create new Request";
                  }
                  default:
                    return "";
                }
              })(item.currStep),
            }))
          );
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
              // style={{ float: "right" }}
            >
              New Request
            </PrimaryButton>
          </div>
          <MarqueeSelection selection={selection}>
            <DetailsList
              items={items}
              columns={columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selection={selection}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
              selectionMode={SelectionMode.single}
            />
          </MarqueeSelection>
          {/* Optionally display selectionDetails somewhere, e.g.: */}
          {/* <div>{selectedId}</div> */}
        </>
      )}
    </div>
  );
};

export default FormsList;
