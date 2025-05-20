import * as React from "react";
import {
  CommandBar,
  ICommandBarItemProps,
} from "@fluentui/react/lib/CommandBar";
import { IButtonProps } from "@fluentui/react/lib/Button";
import { useNavigate } from "react-router-dom";

interface IIndividualCommandBarButtonAsExampleProps {
  onDismissCoachmark: () => void;
  isCoachmarkVisible: boolean;
}

const overflowButtonProps: IButtonProps = {
  ariaLabel: "More commands",
};

/** Command bar which renders the Share button with a coachmark */
const IndividualCommandBarButtonAsExample: React.FunctionComponent<
  IIndividualCommandBarButtonAsExampleProps
> = (props) => {
  const { onDismissCoachmark, isCoachmarkVisible } = props;

  const navigate = useNavigate();

  const items: ICommandBarItemProps[] = React.useMemo(() => {
    return [
      {
        key: "newItem",
        text: "New",
        iconProps: { iconName: "Add" },
        onClick: () => navigate("/form"),
      },
      {
        key: "edit",
        text: "Edit",
        iconProps: { iconName: "Edit" },
        onClick: () => console.log("Upload"),
      },
      {
        key: "share",
        text: "Share",
        iconProps: { iconName: "Share" },
        onClick: () => console.log("Share"),
      },
      {
        key: "download",
        text: "Download",
        iconProps: { iconName: "Download" },
        onClick: () => console.log("Download"),
      },
    ];
  }, [onDismissCoachmark, isCoachmarkVisible]);

  return <CommandBar overflowButtonProps={overflowButtonProps} items={items} />;
};

export const ActionBar: React.FunctionComponent = () => {
  const [isCoachmarkVisible, setIsCoachmarkVisible] = React.useState(true);

  const onDismissCoachmark = React.useCallback(
    () => setIsCoachmarkVisible(false),
    []
  );

  return (
    <IndividualCommandBarButtonAsExample
      onDismissCoachmark={onDismissCoachmark}
      isCoachmarkVisible={isCoachmarkVisible}
    />
  );
};

export default ActionBar;
