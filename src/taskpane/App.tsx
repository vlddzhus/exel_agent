import React, { useState, useEffect } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { Header } from "./components/Header/Header";
import { ChatPanel } from "./components/ChatPanel/ChatPanel";
import { Onboarding } from "./components/Onboarding/Onboarding";
import { Paywall } from "./components/Paywall/Paywall";
import { Settings } from "./components/Settings/Settings";
import { History } from "./components/History/History";
import { useUserStore } from "./stores/userStore";
import { ErrorBoundary } from "./components/common/ErrorBoundary";
import { ToastProvider } from "./components/common/ToastHost";

type View = "onboarding" | "chat" | "paywall" | "settings" | "history";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
    fontFamily: '"Segoe UI", -apple-system, system-ui, sans-serif',
  },
  main: {
    flex: 1,
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
  },
});

export const App: React.FC = () => {
  const styles = useStyles();
  const [view, setView] = useState<View>("onboarding");
  const tier = useUserStore((s) => s.tier);
  const usage = useUserStore((s) => s.usage);

  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        const seen = localStorage.getItem("onboarding_seen");
        if (seen) setView("chat");
      }
    });
  }, []);

  useEffect(() => {
    if (tier === "free" && usage >= 10) {
      setView("paywall");
    }
  }, [tier, usage]);

  const renderView = () => {
    switch (view) {
      case "onboarding":
        return (
          <Onboarding
            onDone={() => {
              localStorage.setItem("onboarding_seen", "1");
              setView("chat");
            }}
          />
        );
      case "chat":
        return <ChatPanel />;
      case "paywall":
        return <Paywall />;
      case "settings":
        return <Settings />;
      case "history":
        return <History onSelectSession={() => setView("chat")} />;
    }
  };

  return (
    <ErrorBoundary>
      <ToastProvider>
        <div className={styles.root}>
          <Header
            onSettingsClick={() => setView("settings")}
            onHistoryClick={() => setView("history")}
            onNewChat={() => setView("chat")}
          />
          <div className={styles.main}>{renderView()}</div>
        </div>
      </ToastProvider>
    </ErrorBoundary>
  );
};
