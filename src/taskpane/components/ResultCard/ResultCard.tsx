import React from "react";
import {
  makeStyles,
  tokens,
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
} from "@fluentui/react-components";
import { useLiveActivityStore } from "../../stores/liveActivityStore";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    margin: "8px 0",
    animation: "crossFade 300ms ease-out",
  },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorStatusSuccessForeground1,
    marginBottom: "8px",
  },
  stats: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    marginBottom: "12px",
    lineHeight: 1.6,
  },
  actions: {
    display: "flex",
    gap: "8px",
  },
});

export const ResultCard: React.FC = () => {
  const styles = useStyles();
  const stats = useLiveActivityStore((s) => s.stats);
  const plan = useLiveActivityStore((s) => s.plan);
  const hasChanges = useLiveActivityStore((s) => s.hasChanges);
  const reset = useLiveActivityStore((s) => s.reset);
  const [undoOpen, setUndoOpen] = React.useState(false);

  const handleRefine = () => {
    const input = document.querySelector("textarea");
    if (input) {
      input.value = "Поправь: ";
      input.focus();
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.title}>✅ {ru.result.done}</div>
      <div className={styles.stats}>
        {plan
          .filter((s) => s.status === "done")
          .map((s) => (
            <div key={s.id}>• {s.label}</div>
          ))}
        <div style={{ marginTop: 4 }}>
          ⏱ {Math.round(stats.elapsedMs / 1000)}с &nbsp; 🪙{" "}
          {stats.tokensOut > 1000
            ? `${(stats.tokensOut / 1000).toFixed(1)}k`
            : stats.tokensOut}{" "}
          токен
        </div>
      </div>
      <div className={styles.actions}>
        {hasChanges && (
          <Dialog open={undoOpen} onOpenChange={(_, d) => setUndoOpen(d.open)}>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" onClick={() => setUndoOpen(true)}>
                {ru.result.undo}
              </Button>
            </DialogTrigger>
            <DialogSurface>
              <DialogBody>
                <DialogTitle>{ru.result.undoConfirm}</DialogTitle>
                <DialogActions>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Отмена</Button>
                  </DialogTrigger>
                  <Button
                    appearance="primary"
                    onClick={() => {
                      setUndoOpen(false);
                      reset();
                    }}
                  >
                    {ru.result.undo}
                  </Button>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        )}
        <Button appearance="subtle" onClick={handleRefine}>
          {ru.result.refine}
        </Button>
      </div>
    </div>
  );
};
