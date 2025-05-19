import * as React from "react";
import styles from "./WorkflowBillApproval.module.scss";
import type { IWorkflowBillApprovalProps } from "./IWorkflowBillApprovalProps";
// import { escape } from "@microsoft/sp-lodash-subset";
import {
  HashRouter,
  Routes,
  Route,
  Navigate,
  useParams,
} from "react-router-dom";
import { ContextStore } from "./Context/ContextStore";
import InitiaterForm from "./Forms/InitiaterForm";

// Define props and state interfaces for the ErrorBoundary
interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
}

class ErrorBoundary extends React.Component<
  ErrorBoundaryProps,
  ErrorBoundaryState
> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false };
  }

  // Update state when an error is thrown
  static getDerivedStateFromError(_: Error): ErrorBoundaryState {
    return { hasError: true };
  }

  // You can log error details here or send to an external service
  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error("Caught by ErrorBoundary:", error, errorInfo);
  }

  render(): React.ReactNode {
    if (this.state.hasError) {
      // Redirect to /err500 on any rendering error
      return <Navigate to="/err500" replace />;
    }

    return this.props.children;
  }
}

// Reusable ErrorPage component
const ErrorPage: React.FC = () => {
  const { code } = useParams();
  let title, message;

  switch (code) {
    case "401":
      title = "401 | Unauthorized";
      message = "You do not have permission to view this page.";
      break;

    case "403":
      title = "403 | Forbidden";
      message = "Access to this resource is denied.";
      break;

    case "429":
      title = "429 | Too Many Requests";
      message = "You have sent too many requests in a short period.";
      break;

    case "404":
      title = "404 | Page Not Found";
      message = "We couldn’t find the page you were looking for.";
      break;

    case "500":
    default:
      title = "500 | Internal Server Error";
      message = "Oops! Something went wrong on our end.";
    // break;

    // default:
    //   title = `${code} | Error`;
    //   message = "An unexpected error occurred.";
  }

  return (
    <div style={{ padding: "2rem", textAlign: "center" }}>
      <h1>{title}</h1>
      <p>{message}</p>
      <div className={styles.links}>
        <a href="#/">Return to Home</a>
      </div>
    </div>
  );
};

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
          <ErrorBoundary>
            <HashRouter>
              <Routes>
                <Route
                  path="/initiaterForm/:formId?"
                  element={<InitiaterForm />}
                />
                {/* <Route path="/gmForm/:formId" element={<GMForm />} /> */}
                <Route path="/err/:code?" element={<ErrorPage />} />
                <Route path="*" element={<Navigate to={"/initiaterForm"} />} />
              </Routes>
            </HashRouter>
          </ErrorBoundary>
        </section>
      </ContextStore.Provider>
    );
  }
}
