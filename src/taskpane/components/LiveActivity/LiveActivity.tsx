import React, { useEffect, useRef } from "react";
import {
  makeStyles,
  tokens,
  ProgressBar,
  Spinner,
  Button,
  Text,
} from "@fluentui/react-components";
import {
  DismissRegular,
  CheckmarkCircleRegular,
  CircleRegular,
  ErrorCircleRegular,
} from "@fluentui/react-icons";
import { useLiveActivityStore } from "../../stores/liveActivityStore";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    margin: "8px 0",
    animation: "slideDown 250ms ease-out",
  },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "8px",
  },
  headerText: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  progress: {
    marginBottom: "8px",
  },
  steps: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    marginBottom: "8px",
  },
  step: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
  },
  stepCurrent: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  stepDone: {
    color: tokens.colorStatusSuccessForeground1,
  },
  thoughts: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    fontStyle: "italic",
    marginBottom: "4px",
    lineHeight: 1.4,
  },
  thought: {
    animation: "fadeInSlideUp 200ms ease-out",
  },
  stats: {
    display: "flex",
    gap: "12px",
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

interface LiveActivityProps {
  onCancel?: () => void;
}

export const LiveActivity: React.FC<LiveActivityProps> = ({ onCancel }) => {
  const styles = useStyles();
  const status = useLiveActivityStore((s) => s.status);
  const plan = useLiveActivityStore((s) => s.plan);
  const thoughts = useLiveActivityStore((s) => s.thoughts);
  const progress = useLiveActivityStore((s) => s.progress);
  const stats = useLiveActivityStore((s) => s.stats);
  const error = useLiveActivityStore((s) => s.error);
  const storeCancel = useLiveActivityStore((s) => s.cancel);
  const thoughtsEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    thoughtsEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [thoughts]);

  const progressValue =
    progress.total > 0 ? progress.done / progress.total : undefined;

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <div className={styles.headerText}>
          {status === "thinking" ? ru.live.thinking : ru.live.plan}
          <span style={{fontSize:10,color:'#999',marginLeft:8}}>v2.1</span>
        </div>
        <Button
          icon={<DismissRegular />}
          appearance="subtle"
          size="small"
          onClick={() => {
            if (onCancel) onCancel();
            storeCancel();
          }}
          aria-label={ru.chat.cancel}
        />
      </div>
      <div className={styles.progress}>
        <ProgressBar value={progressValue} thickness="medium" />
      </div>
      {plan.length > 0 && (
        <div className={styles.steps}>
          {plan.map((step) => (
            <div
              key={step.id}
              className={`${styles.step} ${step.status === "current" ? styles.stepCurrent : ""} ${step.status === "done" ? styles.stepDone : ""}`}
            >
              {step.status === "done" ? (
                <CheckmarkCircleRegular fontSize={14} />
              ) : step.status === "current" ? (
                <Spinner size="tiny" />
              ) : (
                <CircleRegular fontSize={14} />
              )}
              {step.label}
              {step.durationMs
                ? ` [${(step.durationMs / 1000).toFixed(1)}с]`
                : ""}
            </div>
          ))}
        </div>
      )}
      {error && (
        <div style={{color:'#d32f2f',fontSize:'11px',padding:'4px 0',wordBreak:'break-all'}}>
          <ErrorCircleRegular fontSize={14} style={{verticalAlign:'middle',marginRight:4}} />
          {error.code}: {error.message}
        </div>
      )}
      {status === "cancelled" && !error && (
        <div style={{color:'#999',fontSize:'11px',padding:'4px 0'}}>Отменено</div>
      )}
      {thoughts.length > 0 && (
        <div className={styles.thoughts}>
          {thoughts.slice(-3).map((t, i) => (
            <div key={i} className={styles.thought}>
              💭 {t.text}
            </div>
          ))}
          <div ref={thoughtsEndRef} />
        </div>
      )}
      <div className={styles.stats}>
        <span>⏱ {Math.round(stats.elapsedMs / 1000)}с</span>
        <span>
          🪙{" "}
          {stats.tokensOut > 1000
            ? `${(stats.tokensOut / 1000).toFixed(1)}k`
            : stats.tokensOut}{" "}
          токен
        </span>
        {stats.provider && <span>⚡ {stats.provider}</span>}
      </div>
    </div>
  );
};
