import { ISpinnerProps, Spinner, SpinnerSize } from "@fluentui/react";
import * as React from "react";

const Loading: React.FC<ISpinnerProps> = (props) => {
  return (
    <div style={{ margin: "auto", maxWidth: "600px" }}>
      <Spinner
        label="Loading..."
        ariaLive="assertive"
        labelPosition="right"
        size={SpinnerSize.large}
        {...(props || {})}
      />
    </div>
  );
};

export default Loading;
