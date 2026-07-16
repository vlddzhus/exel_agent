import React, { useState, useCallback, createContext, useContext } from "react";
import { makeStyles, tokens, Button, Text } from "@fluentui/react-components";
import {
  CheckmarkCircleRegular,
  ErrorCircleRegular,
  InfoRegular,
  DismissRegular,
} from "@fluentui/react-icons";

type ToastKind = "success" | "error" | "info";

interface Toast {
  id: number;
  kind: ToastKind;
  message: string;
}

interface ToastContextValue {
  showToast: (kind: ToastKind, message: string) => void;
}

const ToastContext = createContext<ToastContextValue>({
  showToast: () => {},
});

export const useToast = () => useContext(ToastContext);

const useStyles = makeStyles({
  host: {
    position: "fixed",
    bottom: "60px",
    left: "12px",
    right: "12px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    zIndex: 2000,
    pointerEvents: "none",
  },
  toast: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow16,
    pointerEvents: "auto",
    animation: "toastSlideIn 250ms ease-out",
  },
  success: {
    backgroundColor: tokens.colorStatusSuccessBackground1,
    color: tokens.colorStatusSuccessForeground1,
  },
  error: {
    backgroundColor: tokens.colorStatusDangerBackground1,
    color: tokens.colorStatusDangerForeground1,
  },
  info: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  message: {
    flex: 1,
    fontSize: tokens.fontSizeBase200,
  },
  icon: { fontSize: "16px" },
});

export function ToastProvider({ children }: { children: React.ReactNode }) {
  const styles = useStyles();
  const [toasts, setToasts] = useState<Toast[]>([]);
  let idCounter = 0;

  const showToast = useCallback((kind: ToastKind, message: string) => {
    const id = ++idCounter;
    setToasts((prev) => [...prev, { id, kind, message }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4000);
  }, []);

  const dismiss = (id: number) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  };

  const kindIcon = (kind: ToastKind) => {
    switch (kind) {
      case "success":
        return <CheckmarkCircleRegular className={styles.icon} />;
      case "error":
        return <ErrorCircleRegular className={styles.icon} />;
      case "info":
        return <InfoRegular className={styles.icon} />;
    }
  };

  return (
    <ToastContext.Provider value={{ showToast }}>
      {children}
      <div className={styles.host}>
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`${styles.toast} ${
              t.kind === "success"
                ? styles.success
                : t.kind === "error"
                  ? styles.error
                  : styles.info
            }`}
          >
            {kindIcon(t.kind)}
            <span className={styles.message}>{t.message}</span>
            <Button
              icon={<DismissRegular />}
              appearance="subtle"
              size="small"
              onClick={() => dismiss(t.id)}
              aria-label="Закрыть"
            />
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}
