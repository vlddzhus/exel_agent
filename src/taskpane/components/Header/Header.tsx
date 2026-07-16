import React from "react";
import {
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
  Badge,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  SettingsRegular,
  HistoryRegular,
  AddRegular,
  MoreHorizontalRegular,
} from "@fluentui/react-icons";
import { useUserStore } from "../../stores/userStore";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    display: "flex",
    alignItems: "center",
    height: "48px",
    padding: "0 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  brand: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    marginRight: "auto",
  },
  spacer: { marginLeft: "auto" },
});

interface HeaderProps {
  onSettingsClick: () => void;
  onHistoryClick: () => void;
  onNewChat: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  onSettingsClick,
  onHistoryClick,
  onNewChat,
}) => {
  const styles = useStyles();
  const tier = useUserStore((s) => s.tier);

  return (
    <div className={styles.root}>
      <span className={styles.brand}>{ru.header.brand}</span>
      {tier !== "free" && (
        <Badge appearance="filled" color="brand" size="small">
          {tier === "pro" ? "Pro" : "Team"}
        </Badge>
      )}
      <Toolbar>
        <ToolbarButton
          icon={<AddRegular />}
          onClick={onNewChat}
          aria-label={ru.header.newChat}
        />
        <ToolbarDivider />
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <ToolbarButton icon={<MoreHorizontalRegular />} aria-label="Меню" />
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<SettingsRegular />} onClick={onSettingsClick}>
                {ru.header.menu.settings}
              </MenuItem>
              <MenuItem icon={<HistoryRegular />} onClick={onHistoryClick}>
                {ru.header.menu.history}
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </Toolbar>
    </div>
  );
};
