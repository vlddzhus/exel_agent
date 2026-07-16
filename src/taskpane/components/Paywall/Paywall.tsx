import React from "react";
import { makeStyles, tokens, Button, Text } from "@fluentui/react-components";
import { LockClosedRegular } from "@fluentui/react-icons";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    padding: "24px",
    textAlign: "center",
    gap: "12px",
  },
  icon: {
    fontSize: "48px",
    color: tokens.colorNeutralForeground3,
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  desc: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    lineHeight: 1.5,
  },
  footnote: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

export const Paywall: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <LockClosedRegular className={styles.icon} />
      <div className={styles.title}>{ru.paywall.title}</div>
      <div className={styles.desc}>{ru.paywall.desc}</div>
      <Button appearance="primary" size="large">
        {ru.paywall.cta}
      </Button>
      <div className={styles.desc}>{ru.paywall.trial}</div>
      <div className={styles.footnote}>{ru.paywall.footnote}</div>
    </div>
  );
};
