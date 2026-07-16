import React from "react";
import { makeStyles, tokens, Button, Text } from "@fluentui/react-components";
import { ErrorCircleRegular } from "@fluentui/react-icons";

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

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
  icon: { fontSize: "32px", color: tokens.colorStatusDangerForeground1 },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  message: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    lineHeight: 1.4,
    wordBreak: "break-word",
  },
});

export class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  ErrorBoundaryState
> {
  state: ErrorBoundaryState = { hasError: false, error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("ErrorBoundary caught:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <ErrorFallback
          message={this.state.error?.message ?? "Произошла ошибка"}
          onRetry={() => this.setState({ hasError: false, error: null })}
        />
      );
    }
    return this.props.children;
  }
}

function ErrorFallback({
  message,
  onRetry,
}: {
  message: string;
  onRetry: () => void;
}) {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <ErrorCircleRegular className={styles.icon} />
      <div className={styles.title}>Что-то пошло не так</div>
      <div className={styles.message}>{message}</div>
      <Button appearance="primary" onClick={onRetry}>
        Попробовать снова
      </Button>
    </div>
  );
}
