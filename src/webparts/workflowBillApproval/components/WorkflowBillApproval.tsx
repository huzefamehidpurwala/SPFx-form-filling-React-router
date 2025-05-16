import * as React from "react";
import styles from "./WorkflowBillApproval.module.scss";
import type { IWorkflowBillApprovalProps } from "./IWorkflowBillApprovalProps";
// import { escape } from "@microsoft/sp-lodash-subset";
import { HashRouter, Routes, Route, Navigate } from "react-router-dom";
import { ContextStore } from "./Context/ContextStore";
import InitiaterForm from "./Forms/InitiaterForm";
import GMForm from "./Forms/GMForm";

export default class WorkflowBillApproval extends React.Component<IWorkflowBillApprovalProps> {
  public render(): React.ReactElement<IWorkflowBillApprovalProps> {
    const { hasTeamsContext, context } = this.props;

    return (
      <ContextStore.Provider value={{ context }}>
        <section
          className={`${styles.workflowBillApproval} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          <HashRouter>
            <Routes>
              <Route path="/initiaterForm" element={<InitiaterForm />} />
              <Route path="/gmForm/:formId" element={<GMForm />} />
              <Route path="/err404" element={<p>404 | Page Not Found</p>} />
              <Route path="/err500" element={<p>500 | Internal Server Error</p>} />
              <Route path="*" element={<Navigate to={"/initiaterForm"} />} />
            </Routes>
          </HashRouter>
        </section>
      </ContextStore.Provider>
    );
  }
}
