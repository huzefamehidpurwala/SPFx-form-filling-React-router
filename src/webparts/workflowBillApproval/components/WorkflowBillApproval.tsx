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
  useNavigate,
} from "react-router-dom";
import { ContextStore } from "./Context/ContextStore";
import Form from "./Forms/Form";
import FormsList from "./Forms/FormsList";
import { PrimaryButton } from "@fluentui/react";

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
      // * not working
      // Redirect to /err/500 on any rendering error
      // return <Navigate to="/err/500" replace />;

      return (
        <div style={{ padding: "2rem", textAlign: "center" }}>
          <h1>{"Oops! Something went wrong on our end."}</h1>
          <PrimaryButton
            iconProps={{ iconName: "Refresh" }}
            onClick={() => window.location.reload()}
          >
            Refresh
          </PrimaryButton>
        </div>
      );
    }

    return this.props.children;
  }
}

// Reusable ErrorPage component
const ErrorPage: React.FC = () => {
  const { code } = useParams();
  const navigate = useNavigate();
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
      <PrimaryButton
        iconProps={{ iconName: "Back" }}
        onClick={() => navigate("/", { replace: true })}
      >
        {"Go to Home"}
      </PrimaryButton>
    </div>
  );
};

const SuccessPage: React.FC = () => {
  const { status } = useParams();
  const navigate = useNavigate();

  const statusMsg: Record<string, string> = {
    app: "Approved",
    rej: "Rejected",
    upd: "Updated",
    sub: "Submitted",
    del: "Deleted",
  };

  return (
    <div style={{ textAlign: "center" }}>
      <p>{statusMsg[status!] || "Done"} successfully</p>
      <PrimaryButton
        iconProps={{ iconName: "Back" }}
        onClick={() => navigate("/", { replace: true })}
      >
        {"Go to Home"}
      </PrimaryButton>
    </div>
  );
};

export default class WorkflowBillApproval extends React.Component<IWorkflowBillApprovalProps> {
  public render(): React.ReactElement<IWorkflowBillApprovalProps> {
    const { hasTeamsContext, context } = this.props;

    return (
      <ContextStore.Provider value={{ spContext: context }}>
        <section
          className={`${styles.workflowBillApproval} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          <HashRouter>
            <ErrorBoundary>
              <Routes>
                <Route path="/form/:formId?" element={<Form />} />
                <Route path="/forms" element={<FormsList />} />
                <Route path="/err/:code" element={<ErrorPage />} />
                <Route path="/succ/:status" element={<SuccessPage />} />
                <Route path="/" element={<Navigate to={"/forms"} replace />} />
                <Route
                  path="*"
                  element={<Navigate to={"/err/404"} replace />}
                />
              </Routes>
            </ErrorBoundary>
          </HashRouter>
        </section>
      </ContextStore.Provider>
    );
  }
}
