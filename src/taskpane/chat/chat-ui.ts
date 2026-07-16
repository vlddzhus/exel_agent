import { renderMarkdown, extractTableAsTsv } from "../utils/markdown";

export function scrollToBottom() {
  const container = document.getElementById("chat-container");
  if (container) {
    requestAnimationFrame(() => {
      container.scrollTop = container.scrollHeight;
    });
  }
}

const labelMap: Record<string, string> = {
  user: "Вы",
  assistant: "AI Агент",
  system: "Система",
  error: "Ошибка",
  "tool-call": "Вызов инструмента",
  "tool-result": "Результат инструмента",
};

// ── Message Actions ──

function getActionButtons(role: string, content: string): string {
  if (role !== "assistant") return "";

  const safeContent = encodeURIComponent(content);
  const hasTable = /<table/i.test(renderMarkdown(content));

  let html = '<div class="message-actions">';

  // Primary actions (always visible, icon-only in narrow)
  html +=
    `<button class="message-actions-btn" data-action="copy" data-content="${safeContent}" title="Copy">` +
    `<span class="action-icon">📋</span><span class="action-label">Copy</span></button>`;

  html +=
    `<button class="message-actions-btn" data-action="retry" data-content="${safeContent}" title="Retry">` +
    `<span class="action-icon">🔄</span><span class="action-label">Retry</span></button>`;

  html +=
    `<button class="message-actions-btn" data-action="edit" data-content="${safeContent}" title="Edit">` +
    `<span class="action-icon">✏️</span><span class="action-label">Edit</span></button>`;

  // Overflow menu for secondary actions
  html +=
    `<div class="msg-actions-overflow">` +
    `<button class="message-actions-btn msg-overflow-toggle" title="More">` +
    `<span class="action-icon">⋯</span></button>` +
    `<div class="msg-overflow-menu" style="display:none">`;

  if (hasTable) {
    html +=
      `<button class="msg-overflow-item" data-action="copy-table" data-content="${safeContent}">` +
      `<span class="action-icon">📊</span> Copy as range</button>`;
  }

  html +=
    `<button class="msg-overflow-item" data-action="save-prompt" data-content="${safeContent}">` +
    `<span class="action-icon">💾</span> Save prompt</button>`;

  html += `</div></div>`;
  html += "</div>";
  return html;
}

let messageCounter = 0;

// ── Main addMessage ──

export function addMessage(
  role: string,
  content: string,
  actionButtons?: { label: string; id: string }[],
) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${role}`;
  msgDiv.dataset.messageId = `msg_${++messageCounter}_${Date.now()}`;

  const label = labelMap[role] || role;
  const rendered =
    role === "user"
      ? escapeHtml(content).replace(/\n/g, "<br>")
      : renderMarkdown(content);

  let html = `<div class="msg-label">${label}</div>
    <div class="message-content-wrapper">
      <div class="message-content">${rendered}</div>
      ${getActionButtons(role, content)}
    </div>`;

  // Tool-call action buttons (backward compat)
  if (actionButtons && actionButtons.length > 0) {
    html += '<div class="msg-actions">';
    for (const btn of actionButtons) {
      html += `<button class="msg-btn" data-action="${btn.id}">${btn.label}</button>`;
    }
    html += "</div>";
  }

  if (role === "assistant") {
    html += `
      <div class="msg-feedback">
        <button class="feedback-btn" data-feedback="helpful" title="Helpful">👍</button>
        <button class="feedback-btn" data-feedback="unhelpful" title="Not helpful">👎</button>
      </div>
    `;
  }

  msgDiv.innerHTML = html;
  messagesDiv.appendChild(msgDiv);

  setupFeedbackListeners(msgDiv);
  setupOverflowMenus(msgDiv);

  scrollToBottom();
}

function setupOverflowMenus(container: HTMLElement) {
  container.addEventListener("click", (e) => {
    const toggle = (e.target as HTMLElement).closest(".msg-overflow-toggle");
    if (!toggle) return;
    e.stopPropagation();

    const overflow = toggle.parentElement?.querySelector(
      ".msg-overflow-menu",
    ) as HTMLElement;
    if (!overflow) return;

    const isVisible = overflow.style.display !== "none";
    // Close all other menus first
    container.querySelectorAll(".msg-overflow-menu").forEach((m) => {
      (m as HTMLElement).style.display = "none";
    });
    overflow.style.display = isVisible ? "none" : "flex";

    if (!isVisible) {
      const close = () => {
        overflow.style.display = "none";
        document.removeEventListener("click", close, true);
      };
      setTimeout(() => document.addEventListener("click", close, true), 0);
    }
  });
}

function setupFeedbackListeners(msgDiv: HTMLElement) {
  const btns = msgDiv.querySelectorAll(".feedback-btn");
  btns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const feedback = (btn as HTMLElement).dataset.feedback;
      const isSelected = btn.classList.contains("active");
      btns.forEach((b) => b.classList.remove("active"));
      if (!isSelected) {
        btn.classList.add("active");
        const event = new CustomEvent("feedback", {
          detail: {
            feedback,
            messageId: msgDiv.dataset.messageId,
            message: msgDiv.textContent?.substring(0, 200),
          },
        });
        document.dispatchEvent(event);
      }
    });
  });
}

export function addPlanCard(
  planText: string,
  onExecute: () => void,
  onCancel: () => void,
) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const card = document.createElement("div");
  card.className = "plan-card";
  card.id = "plan-card";

  const stepsHtml = renderMarkdown(planText);

  card.innerHTML = `
    <div class="plan-header">📋 План действий</div>
    <div class="plan-body">${stepsHtml}</div>
    <div class="plan-actions">
      <button class="plan-btn plan-btn-execute" id="plan-execute">▶ Выполнить всё</button>
      <button class="plan-btn plan-btn-cancel" id="plan-cancel">✕ Отмена</button>
    </div>
  `;

  messagesDiv.appendChild(card);
  scrollToBottom();

  card.querySelector("#plan-execute")?.addEventListener("click", () => {
    card.remove();
    onExecute();
  });

  card.querySelector("#plan-cancel")?.addEventListener("click", () => {
    card.remove();
    onCancel();
  });
}

export function addProgressIndicator() {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const container = document.createElement("div");
  container.className = "progress-container";
  container.id = "progress-container";

  const track = document.createElement("div");
  track.className = "progress-track";

  const bar = document.createElement("div");
  bar.className = "progress-bar";
  bar.id = "progress-bar";
  bar.style.width = "0%";

  track.appendChild(bar);

  container.innerHTML = `
    <div class="progress-label" id="progress-label">Выполняется...</div>
  `;
  container.appendChild(track);

  messagesDiv.appendChild(container);
  scrollToBottom();
}

export function updateProgress(
  current: number,
  total: number,
  label: string,
  elapsed?: number,
) {
  const bar = document.getElementById("progress-bar");
  const labelEl = document.getElementById("progress-label");
  const elapsedEl = document.getElementById("progress-elapsed");
  if (bar) {
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;
    bar.style.width = `${pct}%`;
  }
  if (labelEl) labelEl.textContent = label;

  if (elapsed !== undefined) {
    if (!elapsedEl) {
      const container = document.getElementById("progress-container");
      if (container) {
        const el = document.createElement("div");
        el.className = "progress-elapsed";
        el.id = "progress-elapsed";
        container.appendChild(el);
      }
    }
    const existing = document.getElementById("progress-elapsed");
    if (existing) {
      existing.textContent = `${elapsed.toFixed(1)}s`;
    }
  } else {
    const existing = document.getElementById("progress-elapsed");
    if (existing) existing.remove();
  }
}

export function removeProgress() {
  const container = document.getElementById("progress-container");
  if (container) container.remove();
}

export function addTypingIndicator(): HTMLElement {
  const messagesDiv = document.getElementById("messages");
  const indicator = document.createElement("div");
  indicator.className = "typing-indicator";
  indicator.id = "typing-indicator";
  indicator.innerHTML = "<span></span><span></span><span></span>";
  messagesDiv?.appendChild(indicator);
  scrollToBottom();
  return indicator;
}

export function removeTypingIndicator() {
  const indicator = document.getElementById("typing-indicator");
  if (indicator) {
    indicator.remove();
  }
}

// ── Streaming Thinking Display ──

export function createThinkingMessage(
  phase: "reasoning" | "executing" | "verifying" | "summarizing" = "reasoning",
): HTMLElement {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) {
    const el = document.createElement("div");
    el.className = "message thinking";
    el.id = "thinking-msg";
    return el;
  }

  const existing = document.getElementById("thinking-msg");
  if (existing) existing.remove();

  const el = document.createElement("div");
  el.className = "message thinking";
  el.id = "thinking-msg";

  const labels: Record<string, string> = {
    reasoning: "🤔 Рассуждение",
    executing: "🔧 Выполнение",
    verifying: "🔍 Проверка",
    summarizing: "📝 Отчёт",
  };

  const defaultTexts: Record<string, string> = {
    reasoning: "Думаю...",
    executing: "Работаю...",
    verifying: "Проверяю...",
    summarizing: "Готовлю отчёт...",
  };

  el.innerHTML = `
    <div class="msg-label thinking-label thinking-phase-${phase}">${labels[phase]}</div>
    <div class="thinking-content" id="thinking-content">${defaultTexts[phase]}</div>
  `;

  messagesDiv.appendChild(el);
  scrollToBottom();
  return el;
}

export function addPhaseSeparator(
  phase: "reasoning" | "executing" | "verifying" | "summarizing",
) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const existing = document.getElementById("phase-separator");
  if (existing) existing.remove();

  const labels: Record<string, string> = {
    reasoning: "🤔 Фаза рассуждения",
    executing: "🔧 Фаза выполнения",
    verifying: "🔍 Фаза проверки",
    summarizing: "📝 Фаза отчёта",
  };

  const sep = document.createElement("div");
  sep.className = `phase-separator phase-${phase}`;
  sep.id = "phase-separator";
  sep.innerHTML = `<span>${labels[phase]}</span>`;
  messagesDiv.appendChild(sep);
  scrollToBottom();
}

export function removePhaseSeparator() {
  const sep = document.getElementById("phase-separator");
  if (sep) sep.remove();
}

export function updateThinkingMessage(
  element: HTMLElement,
  text: string,
  phase?: "reasoning" | "executing" | "verifying" | "summarizing",
) {
  const contentEl =
    (element.querySelector("#thinking-content") as HTMLElement) ??
    (element.querySelector(".thinking-content") as HTMLElement);
  if (contentEl) {
    contentEl.textContent = text;
    if (!contentEl.querySelector(".thinking-cursor")) {
      const cursor = document.createElement("span");
      cursor.className = "thinking-cursor";
      contentEl.appendChild(cursor);
    }
    scrollToBottom();
  }

  if (phase) {
    const labelEl = element.querySelector(".thinking-label") as HTMLElement;
    if (labelEl) {
      const labels: Record<string, string> = {
        reasoning: "🤔 Рассуждение",
        executing: "🔧 Выполнение",
        verifying: "🔍 Проверка",
        summarizing: "📝 Отчёт",
      };
      labelEl.textContent = labels[phase];
    }
  }
}

export function finalizeThinkingMessage(
  element: HTMLElement,
  fullText: string,
) {
  const contentEl =
    (element.querySelector("#thinking-content") as HTMLElement) ??
    (element.querySelector(".thinking-content") as HTMLElement);
  if (contentEl) {
    const cursor = contentEl.querySelector(".thinking-cursor");
    if (cursor) cursor.remove();
    contentEl.textContent = fullText;
  }
  element.className = "message assistant";
  const labelEl = element.querySelector(".msg-label");
  if (labelEl) labelEl.textContent = "AI Агент";
  scrollToBottom();
}

// ── Categorized Error Display ──

export function addApiError(message: string, retryInfo?: string) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const el = document.createElement("div");
  el.className = "message api-error";

  let html = `<div class="error-badge">Ошибка API</div><div>${escapeHtml(message)}</div>`;
  if (retryInfo) {
    html += `<button class="error-retry" data-action="retry">${escapeHtml(retryInfo)}</button>`;
  }

  el.innerHTML = html;
  messagesDiv.appendChild(el);
  scrollToBottom();
}

export function addOfficeJsError(message: string, hint?: string) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const el = document.createElement("div");
  el.className = "message office-error";

  let html = `<div class="error-badge">Ошибка Office.js</div><div>${escapeHtml(message)}</div>`;
  if (hint) {
    html += `<div class="error-hint">💡 ${escapeHtml(hint)}</div>`;
  }

  el.innerHTML = html;
  messagesDiv.appendChild(el);
  scrollToBottom();
}

export function addFormulaError(message: string, suggestion?: string) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const el = document.createElement("div");
  el.className = "message formula-error";

  let html = `<div class="error-badge">Ошибка формулы</div><div>${escapeHtml(message)}</div>`;
  if (suggestion) {
    html += `<div class="error-suggestion">→ ${escapeHtml(suggestion)}</div>`;
  }

  el.innerHTML = html;
  messagesDiv.appendChild(el);
  scrollToBottom();
}

export function showConfirmDialog(
  title: string,
  description: string,
): Promise<boolean> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "confirm-overlay";

    const isLongContent =
      description.length > 200 || description.includes("\n");
    const displayContent = isLongContent
      ? `<pre class="confirm-description">${escapeHtml(description)}</pre>`
      : `<p>${escapeHtml(description)}</p>`;

    const box = document.createElement("div");
    box.className = `confirm-box${isLongContent ? " confirm-box-scroll" : ""}`;
    box.innerHTML = `
      <h3>${escapeHtml(title)}</h3>
      ${displayContent}
      <div class="actions">
        <button class="btn-cancel" id="confirm-cancel">Отмена</button>
        <button class="btn-confirm" id="confirm-ok">Подтвердить</button>
      </div>
    `;

    overlay.appendChild(box);
    document.body.appendChild(overlay);

    box.querySelector("#confirm-ok")?.addEventListener("click", () => {
      overlay.remove();
      resolve(true);
    });

    box.querySelector("#confirm-cancel")?.addEventListener("click", () => {
      overlay.remove();
      resolve(false);
    });
  });
}

export function updateConnectionStatus(
  status: "online" | "offline" | "loading",
  text: string,
) {
  const indicator = document.getElementById("status-indicator");
  const statusText = document.getElementById("status-text");
  if (indicator) {
    indicator.className = `indicator ${status}`;
  }
  if (statusText) {
    statusText.textContent = text;
  }
}

export function setInputEnabled(enabled: boolean) {
  const input = document.getElementById("chat-input") as HTMLTextAreaElement;
  const sendBtn = document.getElementById("send-btn") as HTMLButtonElement;
  if (input) input.disabled = !enabled;
  if (sendBtn) sendBtn.disabled = !enabled;
  if (enabled && input) input.focus();
}

// ── Summary Card ──

export interface SummaryItem {
  label: string;
  value: string;
  icon?: string;
}

export function addSummaryCard(items: SummaryItem[], footer?: string) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  const card = document.createElement("div");
  card.className = "summary-card";
  card.id = "summary-card";

  let html =
    '<div class="summary-header">✅ Сводка</div><div class="summary-body">';
  for (const item of items) {
    const icon = item.icon || "•";
    html += `<div class="summary-row"><span class="summary-icon">${icon}</span><span class="summary-label">${escapeHtml(item.label)}</span><span class="summary-value">${escapeHtml(item.value)}</span></div>`;
  }
  html += "</div>";
  if (footer) {
    html += `<div class="summary-footer">${escapeHtml(footer)}</div>`;
  }
  card.innerHTML = html;

  messagesDiv.appendChild(card);
  scrollToBottom();
}

function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}
