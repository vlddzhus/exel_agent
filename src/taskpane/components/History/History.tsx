import React, { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Input,
  Text,
  Button,
  Spinner,
} from "@fluentui/react-components";
import {
  SearchRegular,
  ArrowLeftRegular,
  DeleteRegular,
} from "@fluentui/react-icons";
import { useSessionStore } from "../../stores/sessionStore";
import type { ChatSession } from "../../utils/session-store";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  search: {
    margin: "8px 12px",
  },
  list: {
    flex: 1,
    overflowY: "auto",
    padding: "0 12px",
  },
  group: {
    marginBottom: "12px",
  },
  groupTitle: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginBottom: "4px",
    textTransform: "uppercase",
    letterSpacing: "0.5px",
  },
  item: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    padding: "8px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: "pointer",
    ":hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  itemBody: {
    flex: 1,
    minWidth: 0,
  },
  itemText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    marginBottom: "2px",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  itemMeta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  deleteBtn: {
    visibility: "hidden",
    flexShrink: 0,
  },
  itemHover: {
    selectors: {
      "&:hover $deleteBtn": { visibility: "visible" },
    },
  },
  empty: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    gap: "8px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    padding: "24px",
    textAlign: "center",
  },
  loading: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
  },
});

function groupSessions(
  sessions: ChatSession[],
): { label: string; items: ChatSession[] }[] {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const yesterday = new Date(today.getTime() - 86400000);
  const weekAgo = new Date(today.getTime() - 6 * 86400000);

  const groups: { label: string; items: ChatSession[] }[] = [
    { label: ru.history.today, items: [] },
    { label: ru.history.yesterday, items: [] },
    { label: ru.history.thisWeek, items: [] },
    { label: ru.history.earlier, items: [] },
  ];

  for (const s of sessions) {
    const d = new Date(s.date);
    if (d >= today) groups[0].items.push(s);
    else if (d >= yesterday) groups[1].items.push(s);
    else if (d >= weekAgo) groups[2].items.push(s);
    else groups[3].items.push(s);
  }

  return groups.filter((g) => g.items.length > 0);
}

function formatDate(dateStr: string): string {
  const d = new Date(dateStr);
  return d.toLocaleString("ru-RU", {
    day: "numeric",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  });
}

interface HistoryProps {
  onSelectSession: () => void;
}

export const History: React.FC<HistoryProps> = ({ onSelectSession }) => {
  const styles = useStyles();
  const [search, setSearch] = useState("");
  const sessions = useSessionStore((s) => s.sessions);
  const loading = useSessionStore((s) => s.loading);
  const load = useSessionStore((s) => s.load);
  const remove = useSessionStore((s) => s.remove);
  const getById = useSessionStore((s) => s.getById);

  useEffect(() => {
    load();
  }, [load]);

  const filtered = search
    ? sessions.filter(
        (s) =>
          s.title.toLowerCase().includes(search.toLowerCase()) ||
          s.preview.toLowerCase().includes(search.toLowerCase()),
      )
    : sessions;

  const groups = groupSessions(filtered);

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button
          icon={<ArrowLeftRegular />}
          appearance="subtle"
          onClick={onSelectSession}
          aria-label="Назад"
        />
        <div className={styles.title}>{ru.history.title}</div>
      </div>
      <Input
        className={styles.search}
        contentBefore={<SearchRegular />}
        placeholder={ru.history.search}
        value={search}
        onChange={(_, d) => setSearch(d.value)}
      />
      {loading ? (
        <div className={styles.loading}>
          <Spinner size="small" />
        </div>
      ) : groups.length === 0 ? (
        <div className={styles.empty}>{ru.history.empty}</div>
      ) : (
        <div className={styles.list}>
          {groups.map((g) => (
            <div key={g.label} className={styles.group}>
              <div className={styles.groupTitle}>{g.label}</div>
              {g.items.map((s) => (
                <div
                  key={s.id}
                  className={styles.item}
                  onClick={() => {
                    const full = getById(s.id);
                    if (full && full.messages?.length) {
                      localStorage.setItem(
                        "restore_session",
                        JSON.stringify(full),
                      );
                    }
                    onSelectSession();
                  }}
                >
                  <div className={styles.itemBody}>
                    <div className={styles.itemText}>{s.title}</div>
                    <div className={styles.itemMeta}>
                      {formatDate(s.date)} · {s.stepCount} шаг
                      {s.stepCount !== 1 ? "а" : ""} ·{" "}
                      {s.tokenCount > 1000
                        ? `${(s.tokenCount / 1000).toFixed(1)}k`
                        : s.tokenCount}{" "}
                      токен
                    </div>
                  </div>
                  <Button
                    className={styles.deleteBtn}
                    icon={<DeleteRegular />}
                    appearance="subtle"
                    size="small"
                    onClick={(e) => {
                      e.stopPropagation();
                      remove(s.id);
                    }}
                    aria-label="Удалить"
                  />
                </div>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
