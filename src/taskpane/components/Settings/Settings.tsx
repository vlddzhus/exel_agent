import React from "react";
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Divider,
} from "@fluentui/react-components";
import { useUserStore } from "../../stores/userStore";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    padding: "16px",
    gap: "16px",
    overflowY: "auto",
    flex: 1,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  row: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  value: {
    color: tokens.colorNeutralForeground1,
  },
});

export const Settings: React.FC = () => {
  const styles = useStyles();
  const email = useUserStore((s) => s.email);
  const tier = useUserStore((s) => s.tier);

  return (
    <div className={styles.root}>
      <div className={styles.section}>
        <div className={styles.sectionTitle}>{ru.settings.account}</div>
        <div className={styles.row}>
          <span>{ru.settings.email}</span>
          <span className={styles.value}>{email || "—"}</span>
        </div>
        <div className={styles.row}>
          <span>{ru.settings.tier}</span>
          <span className={styles.value}>
            {tier === "pro" ? "Pro" : tier === "team" ? "Team" : "Free"}
          </span>
        </div>
        <Button appearance="subtle" size="small">
          {ru.settings.manageSubscription}
        </Button>
      </div>
      <Divider />
      <div className={styles.section}>
        <div className={styles.sectionTitle}>{ru.settings.session}</div>
        <Button appearance="subtle" size="small">
          {ru.settings.clearHistory}
        </Button>
        <Button appearance="subtle" size="small">
          {ru.settings.logout}
        </Button>
      </div>
      <Divider />
      <div className={styles.section}>
        <div className={styles.sectionTitle}>{ru.settings.about}</div>
        <div className={styles.row}>
          <span>{ru.settings.version}</span>
          <span className={styles.value}>1.0.0</span>
        </div>
        <Button appearance="subtle" size="small">
          {ru.settings.support}
        </Button>
      </div>
    </div>
  );
};
