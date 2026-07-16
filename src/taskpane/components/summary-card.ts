export interface SummaryStats {
  rowsModified: number;
  actionsExecuted: number;
  timeElapsed: number;
  tokensUsed: number;
  actionTypes?: string[];
}

export interface SuggestionChip {
  icon: string;
  text: string;
}

export interface SummaryCardOptions {
  stats: SummaryStats;
  footer?: string;
  suggestions?: SuggestionChip[];
  onSuggestionClick?: (text: string) => void;
  onAction?: (action: string) => void;
}

export interface SummaryCardAPI {
  element: HTMLElement;
  destroy(): void;
}

export function createSummaryCard(options: SummaryCardOptions): SummaryCardAPI {
  const card = document.createElement("div");
  card.className = "summary-card";
  card.id = "summary-card";

  const { stats } = options;

  const timeStr =
    stats.timeElapsed < 60
      ? `${stats.timeElapsed.toFixed(1)}s`
      : `${Math.floor(stats.timeElapsed / 60)}m ${Math.round(stats.timeElapsed % 60)}s`;

  const actionTypes = stats.actionTypes?.length
    ? stats.actionTypes.slice(0, 3).join(", ") +
      (stats.actionTypes.length > 3 ? "..." : "")
    : "";

  let html = '<div class="summary-header">📋 Done</div>';
  html += '<div class="summary-body">';

  const rows = [
    {
      icon: "📝",
      label: "Rows modified",
      value: formatNumber(stats.rowsModified),
    },
    { icon: "⚡", label: "Actions", value: `${stats.actionsExecuted}` },
    { icon: "⏱", label: "Time", value: timeStr },
    { icon: "🎯", label: "Tokens", value: formatNumber(stats.tokensUsed) },
  ];

  for (const row of rows) {
    html += `<div class="summary-row">
      <span class="summary-icon">${row.icon}</span>
      <span class="summary-label">${escapeHtml(row.label)}</span>
      <span class="summary-value">${escapeHtml(row.value)}</span>
    </div>`;
  }

  if (actionTypes) {
    html += `<div class="summary-row summary-row-actions">
      <span class="summary-icon">🔧</span>
      <span class="summary-label">Tools</span>
      <span class="summary-value summary-value-actions">${escapeHtml(actionTypes)}</span>
    </div>`;
  }

  html += "</div>";

  // Quick action buttons
  html += '<div class="summary-actions">';
  html += `<button class="summary-action-btn" data-action="undo-all">↩ Undo</button>`;
  html += `<button class="summary-action-btn" data-action="refine">✏️ Refine</button>`;
  html += `<button class="summary-action-btn" data-action="export">📤 Export</button>`;
  html += "</div>";

  if (options.footer) {
    html += `<div class="summary-footer">${escapeHtml(options.footer)}</div>`;
  }

  // Smart Suggestions
  if (options.suggestions && options.suggestions.length > 0) {
    html += '<div class="smart-suggestions">';
    for (const chip of options.suggestions) {
      html += `<button class="smart-suggestion-chip" data-prompt="${escapeAttr(chip.text)}">
        <span class="chip-icon">${escapeHtml(chip.icon)}</span>
        <span>${escapeHtml(chip.text)}</span>
      </button>`;
    }
    html += "</div>";
  }

  card.innerHTML = html;

  // Event delegation
  card.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;

    const chip = target.closest(".smart-suggestion-chip") as HTMLElement;
    if (chip) {
      const prompt = chip.dataset.prompt;
      if (prompt) {
        options.onSuggestionClick?.(prompt);
      }
      return;
    }

    const actionBtn = target.closest(".summary-action-btn") as HTMLElement;
    if (actionBtn) {
      const action = actionBtn.dataset.action;
      if (action) {
        options.onAction?.(action);
      }
    }
  });

  return {
    element: card,
    destroy() {
      card.remove();
    },
  };
}

function formatNumber(n: number): string {
  if (n >= 1000000) return (n / 1000000).toFixed(1) + "M";
  if (n >= 1000) return (n / 1000).toFixed(1) + "k";
  return n.toLocaleString();
}

function escapeHtml(s: string): string {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
