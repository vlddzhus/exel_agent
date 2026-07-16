import React, { useRef, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Text,
} from "@fluentui/react-components";
import { SendRegular } from "@fluentui/react-icons";
import { useChatStore } from "../../stores/chatStore";
import { useLiveActivityStore } from "../../stores/liveActivityStore";
import { useAgent } from "../../hooks/useAgent";
import { LiveActivity } from "../LiveActivity/LiveActivity";
import { ResultCard } from "../ResultCard/ResultCard";
import { renderMarkdown } from "../../utils/markdown";
import ru from "../../i18n/ru.json";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
  },
  messages: {
    flex: 1,
    overflowY: "auto",
    padding: "12px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  empty: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    gap: "12px",
    padding: "24px",
    textAlign: "center",
  },
  emptyTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  emptySubtitle: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  chips: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    justifyContent: "center",
  },
  chip: {
    padding: "8px 16px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    cursor: "pointer",
    fontSize: tokens.fontSizeBase200,
    ":hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  inputBar: {
    display: "flex",
    gap: "8px",
    padding: "12px",
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  textarea: {
    flex: 1,
  },
  bubble: {
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
    maxWidth: "85%",
    wordBreak: "break-word",
    fontSize: tokens.fontSizeBase200,
    lineHeight: 1.4,
  },
  userBubble: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    alignSelf: "flex-end",
  },
  assistantBubble: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    alignSelf: "flex-start",
  },
  systemBubble: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    alignSelf: "center",
    fontSize: tokens.fontSizeBase100,
    textAlign: "center",
  },
  errorBubble: {
    backgroundColor: tokens.colorStatusDangerBackground1,
    color: tokens.colorStatusDangerForeground1,
    alignSelf: "center",
  },
});

export const ChatPanel: React.FC = () => {
  const styles = useStyles();
  const messages = useChatStore((s) => s.messages);
  const isProcessing = useChatStore((s) => s.isProcessing);
  const loadMessages = useChatStore((s) => s.loadMessages);
  const liveStatus = useLiveActivityStore((s) => s.status);
  const { sendMessage, cancel } = useAgent();
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const [input, setInput] = React.useState("");

  useEffect(() => {
    const stored = localStorage.getItem("restore_session");
    if (stored) {
      try {
        const session = JSON.parse(stored);
        if (session.messages?.length) {
          loadMessages(session.messages);
        }
      } catch {
        /* ignore */
      }
      localStorage.removeItem("restore_session");
    }
  }, [loadMessages]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, liveStatus]);

  const handleSend = () => {
    const text = input.trim();
    if (!text || isProcessing) return;
    setInput("");
    sendMessage(text);
  };

  const bubbleClass = (role: string) => {
    switch (role) {
      case "user":
        return `${styles.bubble} ${styles.userBubble}`;
      case "assistant":
        return `${styles.bubble} ${styles.assistantBubble}`;
      case "system":
        return `${styles.bubble} ${styles.systemBubble}`;
      case "error":
        return `${styles.bubble} ${styles.errorBubble}`;
      default:
        return styles.bubble;
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.messages} role="log" aria-live="polite">
        {messages.length === 0 && liveStatus === "idle" ? (
          <div className={styles.empty}>
            <div className={styles.emptyTitle}>{ru.chat.emptyTitle}</div>
            <div className={styles.emptySubtitle}>{ru.chat.emptySubtitle}</div>
            <div className={styles.chips}>
              {Object.values(ru.chat.examples).map((ex) => (
                <div
                  key={ex}
                  className={styles.chip}
                  onClick={() => sendMessage(ex)}
                >
                  {ex}
                </div>
              ))}
            </div>
          </div>
        ) : (
          <>
            {messages.map((msg) => (
              <div key={msg.id} className={bubbleClass(msg.role)}>
                {msg.role === "assistant" ? (
                  <div
                    className="message-content"
                    dangerouslySetInnerHTML={{
                      __html: renderMarkdown(msg.content),
                    }}
                  />
                ) : (
                  msg.content
                )}
              </div>
            ))}
            {liveStatus !== "idle" && liveStatus !== "done" && <LiveActivity onCancel={cancel} />}
            {liveStatus === "done" && <ResultCard />}
          </>
        )}
        <div ref={messagesEndRef} />
      </div>
      <div className={styles.inputBar}>
        <Textarea
          className={styles.textarea}
          placeholder={ru.chat.placeholder}
          value={input}
          onChange={(_, d) => setInput(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              handleSend();
            }
          }}
          disabled={isProcessing}
          resize="none"
          rows={1}
        />
        <Button
          icon={<SendRegular />}
          appearance="primary"
          onClick={handleSend}
          disabled={isProcessing || !input.trim()}
          aria-label={ru.chat.send}
        />
      </div>
    </div>
  );
};
