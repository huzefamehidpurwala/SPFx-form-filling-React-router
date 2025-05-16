import * as React from "react";
import styles from "./WorkflowBillApproval.module.scss";
import type { IWorkflowBillApprovalProps } from "./IWorkflowBillApprovalProps";
// import { escape } from "@microsoft/sp-lodash-subset";
import { HashRouter, Routes, Route, useNavigate } from "react-router-dom";

function App() {
  const navigate = useNavigate();
  return (
    <div>
      <p>i am in app comp</p>
      {/* <a href="#/about">about</a> */}
      <button type="button" onClick={() => navigate("/about")}>
        about
      </button>
    </div>
  );
}

function About() {
  return (
    <div>
      <a href="#/">home</a>
      <p>i am in about comp</p>
    </div>
  );
}

export default class WorkflowBillApproval extends React.Component<IWorkflowBillApprovalProps> {
  public render(): React.ReactElement<IWorkflowBillApprovalProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section
        className={`${styles.workflowBillApproval} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <HashRouter>
          <Routes>
            <Route path="/" element={<App />} />
            <Route path="/about" element={<About />} />
          </Routes>
        </HashRouter>
      </section>
    );
  }
}
