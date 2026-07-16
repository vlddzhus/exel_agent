import "./styles/variables.css";
import "./styles/base.css";
import "./styles/dark.css";
import "./styles/header.css";
import "./styles/tabs.css";
import "./styles/welcome.css";
import "./styles/input.css";
import "./styles/plan-card.css";
import "./styles/toast.css";
import "./styles/history.css";
import "./styles/prompts.css";
import "./styles/narrow.css";
import "./styles/markdown.css";
import "./styles/summary-card.css";
import "./taskpane.css";

import { LlmClient, MODEL_CHAIN } from "./chat/llm-client";
import {
  ReActLoop,
  ConfirmationHandler,
  ContinueHandler,
  PlanConfirmHandler,
  ThinkingCallback,
  ExecutionStats,
  detectIntent,
} from "./agent/react-loop";
import {
  addMessage,
  addTypingIndicator,
  removeTypingIndicator,
  showConfirmDialog,
  updateConnectionStatus,
  setInputEnabled,
  addProgressIndicator,
  updateProgress,
  removeProgress,
  createThinkingMessage,
  updateThinkingMessage,
  finalizeThinkingMessage,
  addApiError,
  addOfficeJsError,
  addFormulaError,
  addSummaryCard,
  addPhaseSeparator,
  removePhaseSeparator,
} from "./chat/chat-ui";
import { renderMarkdown, extractTableAsTsv } from "./utils/markdown";
import {
  createSummaryCard,
  SummaryStats,
  SuggestionChip,
} from "./components/summary-card";
import { undoManager } from "./tools/backup";
import "./tools/write";
import "./tools/format";
import "./tools/structure";
import "./tools/transform";
import "./tools/table-tools";
import "./tools/chart-tools";
import "./tools/pivot-tools";
import "./tools/formula-tools";
import "./tools/worksheet-tools";
import "./tools/read";
import "./tools/knowledge-tools";

import { createHeader, HeaderAPI } from "./components/header";
import { createTabBar, TabBarAPI, TabDefinition } from "./components/tab-bar";
import {
  createWelcomeScreen,
  WelcomeScreenAPI,
} from "./components/welcome-screen";
import {
  createMessageInput,
  MessageInputAPI,
} from "./components/message-input";
import {
  createPlanCard,
  PlanCardAPI,
  parsePlanSteps,
  PlanStep,
} from "./components/plan-card";
import { showToast } from "./components/toast";
import { createHistory, HistoryAPI } from "./components/history";
import {
  createPromptLibrary,
  PromptLibraryAPI,
} from "./components/prompt-library";
import {
  ChatSession,
  ChatMessageData,
  getSessions,
  saveSession,
  saveSessionAsync,
  loadSessionAsync,
  deleteSession,
  generateSessionId,
  countTokens,
} from "./utils/session-store";

const BACKEND_URL_STORAGE_KEY = "backend_url";
const MODEL_STORAGE_KEY = "selected_model";
const THEME_STORAGE_KEY = "app_theme";
const CURRENT_SESSION_KEY = "current_session_id";
// HTTPS обязательно: taskpane грузится с https://localhost:3000, и HTTP-fetch
// заблокируется как mixed content (NETWORK_ERROR: Failed to fetch в Office WebView).
// См. backend/.env: USE_HTTPS=true.
const DEFAULT_BACKEND_URL = "https://localhost:4000";
const DEFAULT_MODEL = "llama-3.3-70b-versatile";

let llmClient!: LlmClient;
let reactLoop: ReActLoop | null = null;
let isProcessing = false;
let thinkingMsgEl: HTMLElement | null = null;

let headerAPI!: HeaderAPI;
let tabBarAPI!: TabBarAPI;
let welcomeScreen: WelcomeScreenAPI | null = null;
let inputAPI!: MessageInputAPI;
let currentPlanCard: PlanCardAPI | null = null;
let historyAPI!: HistoryAPI;
let promptLibraryAPI!: PromptLibraryAPI;

let currentTab = "chat";
let currentSessionId = "";

// ── Theme Management ──

function initTheme() {
  const saved = localStorage.getItem(THEME_STORAGE_KEY);
  if (saved === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
  } else if (saved === "light") {
    document.documentElement.removeAttribute("data-theme");
  } else {
    const prefersDark = window.matchMedia(
      "(prefers-color-scheme: dark)",
    ).matches;
    if (prefersDark) {
      document.documentElement.setAttribute("data-theme", "dark");
    }
  }
}

function getTheme(): "light" | "dark" {
  return document.documentElement.getAttribute("data-theme") === "dark"
    ? "dark"
    : "light";
}

// ── Session Management ──

function initSession() {
  currentSessionId =
    localStorage.getItem(CURRENT_SESSION_KEY) || generateSessionId();
  localStorage.setItem(CURRENT_SESSION_KEY, currentSessionId);
}

function getChatMessages(): ChatMessageData[] {
  if (!reactLoop) return [];
  return reactLoop.getMessages().map((m) => ({
    role: m.role,
    content: m.content || "",
  }));
}

function saveCurrentSession() {
  const messages = getChatMessages();
  if (messages.length <= 1) return; // only system prompt

  const firstUserMsg = messages.find((m) => m.role === "user");
  const preview = firstUserMsg?.content?.substring(0, 120) || "Пустой диалог";

  const existing = getSessions().find((s) => s.id === currentSessionId);
  const session: ChatSession = {
    id: currentSessionId,
    title: existing?.title || preview,
    date: new Date().toISOString(),
    preview,
    stepCount: messages.filter((m) => m.role === "tool").length,
    tokenCount: messages.reduce(
      (sum, m) => sum + countTokens(m.content || ""),
      0,
    ),
    messages,
  };
  saveSession(session); // sync to localStorage
  saveSessionAsync(session).catch(() => {}); // async to backend
  if (historyAPI) historyAPI.refresh();
}

async function loadSession(session: ChatSession) {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;

  // Try to get full messages from server if this is a meta-only object
  let fullSession = session;
  if (!session.messages || session.messages.length === 0) {
    const serverSession = await loadSessionAsync(session.id);
    if (serverSession && serverSession.messages) {
      fullSession = serverSession;
    }
  }

  messagesDiv.innerHTML = "";
  if (welcomeScreen) {
    welcomeScreen.destroy();
    welcomeScreen = null;
  }

  currentSessionId = fullSession.id;
  localStorage.setItem(CURRENT_SESSION_KEY, currentSessionId);

  if (reactLoop) {
    reactLoop.clearConversation();
  }

  for (const msg of fullSession.messages) {
    if (msg.role === "system") {
      addMessage("system", msg.content);
    } else if (msg.role === "user") {
      addMessage("user", msg.content);
    } else if (msg.role === "assistant") {
      addMessage("assistant", msg.content);
    }
  }

  switchTab("chat");
  showToast({ message: "Сессия восстановлена", type: "info" });
}

// ── Tab Management ──

const TABS: TabDefinition[] = [
  { id: "chat", label: "Чат", icon: "💬" },
  { id: "history", label: "История", icon: "📋", badge: 0 },
  { id: "prompts", label: "Запросы", icon: "⚡" },
];

function switchTab(tabId: string) {
  currentTab = tabId;
  tabBarAPI.setActiveTab(tabId);

  const chatContainer = document.getElementById("chat-container");
  const inputArea = document.getElementById("input-area");
  const panels = document.querySelectorAll<HTMLElement>(".tab-panel");

  if (tabId === "chat") {
    if (chatContainer) chatContainer.style.display = "flex";
    if (inputArea) inputArea.style.display = "flex";
    panels.forEach((p) => p.classList.remove("active"));
  } else {
    if (chatContainer) chatContainer.style.display = "none";
    if (inputArea) inputArea.style.display = "none";
    panels.forEach((p) => p.classList.remove("active"));
    const panelId = `panel-${tabId}`;
    const panel = document.getElementById(panelId);
    if (panel) panel.classList.add("active");

    if (tabId === "history" && historyAPI) {
      historyAPI.refresh();
    }
  }

  if (tabId === "chat" && inputArea) {
    inputArea.style.display = "flex";
  }
}

// ── Welcome Screen ──

function showWelcome() {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return;
  const existing = document.getElementById("welcome-screen");
  if (existing) existing.remove();

  welcomeScreen = createWelcomeScreen({
    onSuggestionClick: (text: string) => {
      handlePromptSuggestion(text);
    },
  });
  messagesDiv.appendChild(welcomeScreen.element);
}

function handlePromptSuggestion(text: string) {
  if (welcomeScreen) {
    welcomeScreen.destroy();
    welcomeScreen = null;
  }
  const input = document.getElementById("chat-input") as HTMLTextAreaElement;
  if (input) {
    input.value = text;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    handleSend();
  }
}

// ── Stop Button ──

function showStopButton() {
  const btn = document.getElementById("stop-btn");
  if (btn) btn.classList.add("visible");
}

function hideStopButton() {
  const btn = document.getElementById("stop-btn");
  if (btn) btn.classList.remove("visible");
}

function setupStopButton() {
  const btn = document.getElementById("stop-btn");
  if (!btn) return;
  btn.addEventListener("click", () => {
    if (reactLoop) {
      reactLoop.abort();
    }
    hideStopButton();
    showToast({ message: "Генерация остановлена", type: "warning" });
  });
}

// ── Current Selection ──

function setupSelectionButton() {
  if (!Office || !Office.context) return;

  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.BindingSelectionChanged,
      updateSelectionDisplay,
    );
  } catch {
    // Selection handler not available
  }

  // Poll on focus to catch selection changes in Excel
  document.addEventListener("focusin", () => {
    setTimeout(updateSelectionDisplay, 300);
  });

  // Poll periodically as a fallback (increased to 10s)
  setInterval(updateSelectionDisplay, 10000);

  updateSelectionDisplay();
}

function updateSelectionDisplay() {
  if (!inputAPI) return;
  try {
    Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      const address = range.address;
      if (address) {
        inputAPI.setSelectionLabel(address);
        inputAPI.setSelectionVisible(true);
      }
    }).catch(() => {
      inputAPI.setSelectionVisible(false);
    });
  } catch {
    inputAPI.setSelectionVisible(false);
  }
}

// ── Backward Compat Helpers ──

function getBackendUrl(): string {
  return localStorage.getItem(BACKEND_URL_STORAGE_KEY) || DEFAULT_BACKEND_URL;
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function showSettingsDialog() {
  const overlay = document.createElement("div");
  overlay.className = "confirm-overlay";

  const currentUrl =
    localStorage.getItem(BACKEND_URL_STORAGE_KEY) || DEFAULT_BACKEND_URL;
  const currentTheme = getTheme();

  overlay.innerHTML = `
    <div class="settings-box" role="dialog" aria-label="Настройки">
      <h3>Настройки</h3>
      <p style="font-size:12px;color:var(--text-secondary);margin:0 0 12px;line-height:1.4">
        URL прокси-сервера. API-ключ хранится на сервере.
      </p>
      <div class="settings-field">
        <label for="settings-url">URL бэкенда</label>
        <input type="text" id="settings-url" value="${escapeAttr(currentUrl)}" placeholder="https://localhost:4000" />
      </div>
      <div class="settings-field">
        <label for="settings-theme">Тема</label>
        <select id="settings-theme" style="display:block;width:100%;padding:10px 16px;border:1px solid var(--text-primary);border-radius:var(--radius-inputs);font:inherit;background:var(--surface-elevated);color:var(--text-primary);outline:none">
          <option value="system" ${currentTheme === "light" && !localStorage.getItem(THEME_STORAGE_KEY) ? "selected" : ""}>Системная</option>
          <option value="light" ${localStorage.getItem(THEME_STORAGE_KEY) === "light" ? "selected" : ""}>Светлая</option>
          <option value="dark" ${localStorage.getItem(THEME_STORAGE_KEY) === "dark" ? "selected" : ""}>Тёмная</option>
        </select>
      </div>
      <div class="actions">
        <button class="btn-cancel" id="settings-cancel">Отмена</button>
        <button class="btn-confirm" id="settings-save">Сохранить</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  const urlInput = overlay.querySelector("#settings-url") as HTMLInputElement;
  const themeSelect = overlay.querySelector(
    "#settings-theme",
  ) as HTMLSelectElement;

  setTimeout(() => urlInput.focus(), 100);

  overlay.querySelector("#settings-save")?.addEventListener("click", () => {
    const url = urlInput.value.trim();
    if (!url) {
      urlInput.style.borderColor = "#d32f2f";
      urlInput.focus();
      return;
    }

    localStorage.setItem(BACKEND_URL_STORAGE_KEY, url);

    const theme = themeSelect.value;
    if (theme === "system") {
      localStorage.removeItem(THEME_STORAGE_KEY);
      const prefersDark = window.matchMedia(
        "(prefers-color-scheme: dark)",
      ).matches;
      if (prefersDark) {
        document.documentElement.setAttribute("data-theme", "dark");
      } else {
        document.documentElement.removeAttribute("data-theme");
      }
    } else if (theme === "dark") {
      localStorage.setItem(THEME_STORAGE_KEY, "dark");
      document.documentElement.setAttribute("data-theme", "dark");
    } else {
      localStorage.setItem(THEME_STORAGE_KEY, "light");
      document.documentElement.removeAttribute("data-theme");
    }

    overlay.remove();

    if (!llmClient) {
      initLlmClient(url);
    } else {
      llmClient.setBackendUrl(url);
    }

    initReActLoop();

    updateConnectionStatus("online", "Подключено");
    addMessage("system", "URL бэкенда сохранён. Готов к работе!");
    setInputEnabled(true);
    showToast({ message: "Настройки сохранены", type: "success" });
  });

  overlay.querySelector("#settings-cancel")?.addEventListener("click", () => {
    overlay.remove();
  });

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });
}

function initLlmClient(backendUrl: string) {
  const savedModel = localStorage.getItem(MODEL_STORAGE_KEY) || DEFAULT_MODEL;
  llmClient = new LlmClient({
    backendUrl,
    model: savedModel,
    temperature: 0.3,
    maxTokens: 8192,
  });

  llmClient.setOnModelFallback((oldModel, newModel) => {
    const modelNames: Record<string, string> = {
      "llama-3.3-70b-versatile": "Llama 3.3 70B",
      "llama-4-scout-17b-16e-instruct": "Llama 4 Scout",
      "qwen-qwen3-32b": "Qwen3 32B",
      "deepseek-r1-distill-70b": "DeepSeek R1 70B",
    };
    const oldName = modelNames[oldModel] || oldModel;
    const newName = modelNames[newModel] || newModel;
    addMessage("system", `⚠️ Модель переключена: ${oldName} → ${newName}`);
    showToast({ message: `Модель переключена на ${newName}`, type: "warning" });
    if (headerAPI) {
      headerAPI.setModel(newModel);
    }
  });
}

function setupModelSelector() {
  const select = document.getElementById("model-select") as HTMLSelectElement;
  if (!select) return;

  const savedModel = localStorage.getItem(MODEL_STORAGE_KEY);
  if (savedModel && MODEL_CHAIN.includes(savedModel)) {
    select.value = savedModel;
  }

  select.addEventListener("change", () => {
    const newModel = select.value;
    localStorage.setItem(MODEL_STORAGE_KEY, newModel);
    if (llmClient) {
      llmClient.setModel(newModel);
    }
  });
}

// ── Callbacks for Agent ──

const onThinking: ThinkingCallback = (
  text: string,
  phase?: "reasoning" | "executing" | "verifying" | "summarizing",
) => {
  if (!thinkingMsgEl || !document.body.contains(thinkingMsgEl)) {
    thinkingMsgEl = createThinkingMessage(phase || "reasoning");
  }
  updateThinkingMessage(thinkingMsgEl, text, phase);
};

function makeConfirmationHandler(): ConfirmationHandler {
  return async (
    toolName: string,
    args: Record<string, unknown>,
    description?: string,
  ) => {
    return showConfirmDialog(
      `Подтвердить: ${toolName}`,
      description || `Аргументы:\n${JSON.stringify(args, null, 2)}`,
    );
  };
}

function makePlanConfirmHandler(): PlanConfirmHandler {
  return async (planText: string, steps: PlanStep[]) => {
    return new Promise((resolve) => {
      if (currentPlanCard) {
        currentPlanCard.destroy();
        currentPlanCard = null;
      }

      currentPlanCard = createPlanCard({
        planText,
        steps,
        onResult: (result) => {
          currentPlanCard = null;
          if (result.approved) {
            showToast({
              message: "План утверждён. Выполняю...",
              type: "success",
            });
          }
          resolve(result);
        },
      });

      const messagesDiv = document.getElementById("messages");
      if (messagesDiv) {
        messagesDiv.appendChild(currentPlanCard.element);
      }
    });
  };
}

function onProgress(current: number, total: number, label: string) {
  updateProgress(current, total, label);
}

// ── Send / Handle Message ──

async function handleSend() {
  if (isProcessing) return;

  const input = document.getElementById("chat-input") as HTMLTextAreaElement;
  const text = input.value.trim();
  if (!text) return;

  if (welcomeScreen) {
    welcomeScreen.destroy();
    welcomeScreen = null;
  }

  inputAPI.clear();

  addMessage("user", text);
  isProcessing = true;
  setInputEnabled(false);
  showStopButton();

  if (!reactLoop) {
    addMessage("error", "Нет подключения. Проверьте URL бэкенда (шестерёнка).");
    isProcessing = false;
    setInputEnabled(true);
    hideStopButton();
    return;
  }

  // ── Intent Detection ──
  const intent = await detectIntent(text, llmClient);
  const isTask = intent === "task";

  if (isTask) {
    // Task mode: plan card → approve → execute
    addProgressIndicator();
    addTypingIndicator();
    thinkingMsgEl = null;

    try {
      await reactLoop.planAndExecute(text);
      thinkingMsgEl = null;
      saveCurrentSession();
    } catch (error: unknown) {
      const msg = error instanceof Error ? error.message : String(error);

      if (msg.includes("cancelled") || msg.includes("abort")) {
        addMessage("system", "Generation cancelled.");
      } else if (
        msg.includes("429") ||
        msg.includes("503") ||
        msg.includes("rate limit") ||
        msg.includes("Rate limit")
      ) {
        addApiError(
          "Достигнут лимит запросов API. Модель переключится автоматически.",
          "Повтор с другой моделью...",
        );
      } else if (
        msg.includes("Office") ||
        msg.includes("Excel") ||
        msg.includes("Range") ||
        msg.includes("context.sync")
      ) {
        addOfficeJsError(
          msg,
          "Попробуйте выбрать другой диапазон или обновить книгу.",
        );
      } else if (
        msg.includes("#NAME?") ||
        msg.includes("Name error") ||
        msg.includes("formula") ||
        msg.includes("Syntax")
      ) {
        addFormulaError(
          msg,
          "Проверьте пропущенные операторы между ссылками на ячейки.",
        );
      } else {
        addMessage("error", `Ошибка: ${msg}`);
        showToast({ message: "Произошла ошибка", type: "error" });
      }
      thinkingMsgEl = null;
    } finally {
      removeTypingIndicator();
      removeProgress();
    }
  } else {
    // Chat mode: simple conversation, no plan, no progress
    thinkingMsgEl = null;
    addTypingIndicator();

    try {
      await reactLoop.handleChat(text);
      saveCurrentSession();
    } catch (error: unknown) {
      const msg = error instanceof Error ? error.message : String(error);
      if (msg.includes("cancelled") || msg.includes("abort")) {
        addMessage("system", "Генерация отменена.");
      } else {
        addMessage("error", `Ошибка: ${msg}`);
        showToast({ message: "Произошла ошибка", type: "error" });
      }
    }
  }

  // Common cleanup
  isProcessing = false;
  setInputEnabled(true);
  hideStopButton();
  removeTypingIndicator();
  removeProgress();
  updateUndoButton();
}

function updateUndoButton() {
  if (headerAPI) {
    const stack = undoManager.getStack();
    headerAPI.setUndoCount(stack.length);
  }
}

async function handleNewChat() {
  saveCurrentSession();

  currentSessionId = generateSessionId();
  localStorage.setItem(CURRENT_SESSION_KEY, currentSessionId);

  const messagesDiv = document.getElementById("messages");
  if (messagesDiv) {
    messagesDiv.innerHTML = "";
  }
  if (reactLoop) {
    reactLoop.clearConversation();
  }
  if (currentPlanCard) {
    currentPlanCard.destroy();
    currentPlanCard = null;
  }
  undoManager.clear();
  showWelcome();
  updateUndoButton();
  showToast({ message: "Новый диалог начат", type: "info" });
}

async function handleUndo() {
  const latest = undoManager.getLatestBackup();
  if (!latest) {
    showToast({ message: "Нет действий для отмены", type: "info" });
    return;
  }

  // Preview: show what will be restored
  const cellInfo = latest.cellCount > 0 ? `${latest.cellCount} ячеек` : "";
  const chunkInfo =
    latest.chunks.length > 1 ? ` (${latest.chunks.length} чанков)` : "";
  const previewLines = [
    `Действие: ${latest.description}`,
    `Диапазон: ${latest.address}`,
    cellInfo ? `Объём: ${cellInfo}${chunkInfo}` : "",
    `Время: ${new Date(latest.timestamp).toLocaleTimeString()}`,
    "",
    "Восстановить эти данные?",
  ]
    .filter(Boolean)
    .join("\n");

  const confirmed = await showConfirmDialog(
    "Отменить последнее действие",
    previewLines,
  );

  if (!confirmed) return;

  const result = await undoManager.restoreBackup(latest.id);
  if (result.success) {
    addMessage(
      "system",
      `✅ Отмена успешна: ${result.description || latest.description}\nДиапазон: ${latest.address}`,
    );
    showToast({ message: "Отмена успешна", type: "success" });
  } else {
    addMessage("error", `Отмена не удалась: ${result.error}`);
    showToast({ message: "Отмена не удалась", type: "error" });
  }

  updateUndoButton();
}

// ── Smart Suggestions Generator ──

function generateSuggestions(
  stats: ExecutionStats,
  userInput: string,
): SuggestionChip[] {
  const chips: SuggestionChip[] = [];
  const toolNames = stats.toolNames.map((t) => t.toLowerCase());
  const input = userInput.toLowerCase();

  // Приоритет подсказок на основе использованных инструментов
  const hasCharts = toolNames.some((t) => /chart|graph|plot/.test(t));
  const hasSortFilter = toolNames.some((t) => /sort|filter|group/.test(t));
  const hasCleanup = toolNames.some((t) =>
    /clean|duplicate|remove|empty|null|missing|trim|replace/.test(t),
  );
  const hasFormat = toolNames.some((t) =>
    /setValues|fillFormula|setRangeFormat|format/.test(t),
  );
  const hasTableOps = toolNames.some((t) =>
    /createTable|addTableRow|sortTable|filterTable/.test(t),
  );
  const hasPivot = toolNames.some((t) => /pivot|crosstab/.test(t));
  const hasFormula = toolNames.some((t) =>
    /formula|calculate|sum|average/.test(t),
  );

  if (hasCharts) {
    chips.push({ icon: "🎨", text: "Изменить стиль или цвета диаграммы" });
    chips.push({ icon: "📊", text: "Добавить подписи данных и заголовок" });
    chips.push({
      icon: "📤",
      text: "Экспортировать диаграмму как изображение",
    });
    chips.push({
      icon: "📋",
      text: "Создать вторую диаграмму из других данных",
    });
  }

  if (hasSortFilter) {
    chips.push({
      icon: "📈",
      text: "Создать диаграмму из отсортированных данных",
    });
    chips.push({ icon: "📝", text: "Суммировать отфильтрованные результаты" });
    chips.push({ icon: "🔍", text: "Найти выбросы в отфильтрованных данных" });
    chips.push({ icon: "📋", text: "Применить условное форматирование" });
  }

  if (hasCleanup) {
    chips.push({ icon: "📊", text: "Создать отчёт по очищенным данным" });
    chips.push({ icon: "🔍", text: "Проверить качество данных после очистки" });
    chips.push({
      icon: "🎨",
      text: "Форматировать очищенные данные как таблицу",
    });
    chips.push({ icon: "📈", text: "Анализировать распределение значений" });
  }

  if (hasFormat || hasTableOps) {
    chips.push({ icon: "📊", text: "Анализировать отформатированные данные" });
    chips.push({ icon: "🧹", text: "Проверить на ошибки и аномалии" });
    chips.push({ icon: "📈", text: "Создать диаграмму из таблицы" });
    chips.push({
      icon: "🔄",
      text: "Применить то же форматирование к другому диапазону",
    });
  }

  if (hasPivot) {
    chips.push({ icon: "📈", text: "Создать диаграмму из сводной таблицы" });
    chips.push({ icon: "🎯", text: "Добавить фильтры или срезы" });
    chips.push({ icon: "🔄", text: "Обновить сводную с новыми данными" });
    chips.push({ icon: "📋", text: "Форматировать макет сводной таблицы" });
  }

  if (hasFormula) {
    chips.push({ icon: "📝", text: "Объяснить использованные формулы" });
    chips.push({ icon: "🎯", text: "Применить формулы к другому диапазону" });
    chips.push({ icon: "📊", text: "Построить диаграмму результатов" });
    chips.push({ icon: "🔍", text: "Проверить результаты формул" });
  }

  // Input-based suggestions (when tool data is thin)
  if (/\b(analy|statistic|mean|avg|total|sum)\b/.test(input) && !hasCharts) {
    chips.push({ icon: "📈", text: "Создать диаграмму на основе анализа" });
  }
  if (/\b(pivot|crosstab|aggregate)\b/.test(input) && !hasPivot) {
    chips.push({ icon: "📊", text: "Построить диаграмму сводной таблицы" });
    chips.push({ icon: "📋", text: "Форматировать макет сводной таблицы" });
  }

  // Always add a few general follow-ups if we have room
  if (stats.rowsModified > 0) {
    chips.push({
      icon: "↩",
      text: `Отменить последние ${stats.rowsModified} строк`,
    });
  }

  // De-duplicate and limit to 5
  const seen = new Set<string>();
  const unique: SuggestionChip[] = [];
  for (const c of chips) {
    if (!seen.has(c.text)) {
      seen.add(c.text);
      unique.push(c);
    }
    if (unique.length >= 5) break;
  }

  if (unique.length < 3) {
    unique.push({ icon: "🔄", text: "Сделать то же для другого диапазона" });
    unique.push({ icon: "📈", text: "Создать диаграмму из результата" });
  }

  return unique;
}

// ── Execution Complete Handler ──

let lastExecutionStats: ExecutionStats | null = null;
let continueResolver: ((value: boolean) => void) | null = null;

function onExecutionComplete(stats: ExecutionStats) {
  lastExecutionStats = stats;

  // Only show summary card for task execution with actual tool calls
  if (stats.toolCallCount > 0) {
    removePhaseSeparator();

    const userInput = getLastUserInput();
    const suggestions = generateSuggestions(stats, userInput);

    const timeStr =
      stats.elapsedMs < 60000
        ? `${(stats.elapsedMs / 1000).toFixed(1)}s`
        : `${Math.floor(stats.elapsedMs / 60000)}m ${Math.round((stats.elapsedMs % 60000) / 1000)}s`;

    const footer = `Задача выполнена за ${timeStr}. Что вы хотите сделать дальше?`;

    const card = createSummaryCard({
      stats: {
        rowsModified: stats.rowsModified,
        actionsExecuted: stats.toolCallCount,
        timeElapsed: stats.elapsedMs / 1000,
        tokensUsed: stats.totalTokens || countApproxTokens(stats.finalMessage),
        actionTypes: stats.toolNames,
      },
      footer,
      suggestions,
      onSuggestionClick: (text: string) => {
        handleSmartSuggestion(text);
      },
      onAction: (action: string) => {
        handleSummaryAction(action, stats);
      },
    });

    const messagesDiv = document.getElementById("messages");
    if (messagesDiv) {
      messagesDiv.appendChild(card.element);
      setTimeout(() => {
        const chatContainer = document.getElementById("chat-container");
        if (chatContainer) {
          chatContainer.scrollTop = chatContainer.scrollHeight;
        }
      }, 50);
    }

    showToast({ message: `✅ Готово за ${timeStr}`, type: "success" });
  }
}

function getLastUserInput(): string {
  const messagesDiv = document.getElementById("messages");
  if (!messagesDiv) return "";
  const userMsgs = messagesDiv.querySelectorAll(".message.user");
  const last = userMsgs[userMsgs.length - 1];
  return last?.textContent?.replace("You", "").trim() || "";
}

function countApproxTokens(text: string): number {
  return Math.ceil((text || "").length / 4);
}

function handleSmartSuggestion(text: string) {
  const input = document.getElementById("chat-input") as HTMLTextAreaElement;
  if (input) {
    input.value = text;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    handleSend();
  }
}

function handleSummaryAction(action: string, stats: ExecutionStats) {
  if (action === "undo-all" && lastExecutionStats) {
    handleUndo();
  } else if (action === "refine") {
    const input = document.getElementById("chat-input") as HTMLTextAreaElement;
    if (input) {
      input.value = "Уточнить предыдущий результат: ";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.focus();
      showToast({ message: "Введите уточнение...", type: "info" });
    }
  } else if (action === "export") {
    const messagesDiv = document.getElementById("messages");
    if (messagesDiv) {
      const text = messagesDiv.textContent || "";
      navigator.clipboard.writeText(text).then(() => {
        showToast({
          message: "Сессия скопирована в буфер обмена",
          type: "success",
        });
      });
    }
  }
}

// ── ReActLoop Factory ──

function initReActLoop() {
  if (!llmClient) return;

  // Wrap onThinking to inject phase separator on transition
  const wrappedThinking: ThinkingCallback = (text, phase) => {
    if (phase === "executing") {
      // Add phase separator when transitioning to execution
      const existing = document.getElementById("phase-separator");
      if (!existing) {
        addPhaseSeparator("executing");
      }
    }
    onThinking(text, phase);
  };

  reactLoop = new ReActLoop(
    llmClient,
    addMessage,
    makeConfirmationHandler(),
    makePlanConfirmHandler(),
    onProgress,
    wrappedThinking,
    onExecutionComplete,
    () => {
      return new Promise<boolean>((resolve) => {
        continueResolver = resolve;
        addMessage(
          "system",
          "🔄 Требуется больше шагов для завершения задачи. Продолжить?",
          [
            { label: "▶ Продолжить", id: "continue_exec" },
            { label: "⏹ Остановить", id: "stop_exec" },
          ],
        );
      });
    },
  );
}

function setupEventListeners() {
  document.addEventListener("click", async (e: MouseEvent) => {
    const target = e.target as HTMLElement;

    const action = target.dataset.action;
    if (action?.startsWith("undo_")) {
      const backupId = action.slice(5);
      const stack = undoManager.getStack();
      const entry = stack.find((e) => e.id === backupId);
      if (!entry) {
        showToast({ message: "Резервная копия не найдена", type: "error" });
        return;
      }
      const previewLines = [
        `Действие: ${entry.description}`,
        `Диапазон: ${entry.address}`,
        entry.cellCount > 0 ? `Объём: ${entry.cellCount} ячеек` : "",
        "",
        "Восстановить эти данные?",
      ]
        .filter(Boolean)
        .join("\n");
      const confirmed = await showConfirmDialog(
        "Отменить действие",
        previewLines,
      );
      if (!confirmed) return;
      const result = await undoManager.restoreBackup(backupId);
      if (result.success) {
        addMessage(
          "system",
          `✅ Отмена успешна: ${result.description || entry.description}`,
        );
        showToast({ message: "Отмена успешна", type: "success" });
      } else {
        addMessage("error", `Отмена не удалась: ${result.error}`);
        showToast({ message: "Отмена не удалась", type: "error" });
      }
      updateUndoButton();
    }

    // Code copy buttons
    if (target.classList.contains("code-copy")) {
      const code = target.dataset.code;
      if (code) {
        try {
          await navigator.clipboard.writeText(code);
          target.textContent = "Скопировано!";
          target.classList.add("copied");
          setTimeout(() => {
            target.textContent = "Копировать";
            target.classList.remove("copied");
          }, 2000);
        } catch {
          showToast({ message: "Не удалось скопировать код", type: "error" });
        }
      }
    }

    // Code insert into sheet
    if (target.classList.contains("code-insert")) {
      const code = target.dataset.code;
      if (code) {
        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("address");
            await context.sync();
            const cell = sheet
              .getRange(range.address)
              .getLastCell()
              .getOffsetRange(0, 1);
            cell.values = [[code]];
            cell.format.autofitColumns();
          });
          showToast({ message: "Код вставлен в лист", type: "success" });
        } catch {
          showToast({
            message: "Не удалось вставить код в лист",
            type: "error",
          });
        }
      }
    }

    // Message actions
    const msgAction = target.dataset.action;

    if (msgAction === "copy") {
      const content = target.dataset.content;
      if (content) {
        try {
          await navigator.clipboard.writeText(decodeURIComponent(content));
          showToast({
            message: "Сообщение скопировано в буфер",
            type: "success",
          });
        } catch {
          showToast({ message: "Не удалось скопировать", type: "error" });
        }
      }
    }

    if (msgAction === "copy-table") {
      const content = target.dataset.content;
      if (content) {
        const tsv = extractTableAsTsv(decodeURIComponent(content));
        if (tsv) {
          try {
            await navigator.clipboard.writeText(tsv);
            showToast({
              message: "Таблица скопирована как диапазон Excel",
              type: "success",
            });
          } catch {
            showToast({
              message: "Не удалось скопировать таблицу",
              type: "error",
            });
          }
        }
      }
    }

    if (msgAction === "retry") {
      const content = target.dataset.content;
      if (content) {
        const text = decodeURIComponent(content);
        const input = document.getElementById(
          "chat-input",
        ) as HTMLTextAreaElement;
        if (input) {
          input.value = text;
          input.dispatchEvent(new Event("input", { bubbles: true }));
        }
        // Remove welcome screen if present
        if (welcomeScreen) {
          welcomeScreen.destroy();
          welcomeScreen = null;
        }
        handleSend();
      }
    }

    if (msgAction === "edit") {
      const content = target.dataset.content;
      if (content) {
        const text = decodeURIComponent(content);
        const input = document.getElementById(
          "chat-input",
        ) as HTMLTextAreaElement;
        if (input) {
          input.value = text;
          input.dispatchEvent(new Event("input", { bubbles: true }));
          input.focus();
          const len = input.value.length;
          input.setSelectionRange(len, len);
        }
        switchTab("chat");
      }
    }

    if (msgAction === "save-prompt") {
      const content = target.dataset.content;
      if (content) {
        const text = decodeURIComponent(content).substring(0, 200);
        try {
          const existing = JSON.parse(
            localStorage.getItem("saved_prompts") || "[]",
          );
          existing.push({ text, date: new Date().toISOString() });
          localStorage.setItem("saved_prompts", JSON.stringify(existing));
          showToast({
            message: "Запрос сохранён в библиотеку",
            type: "success",
          });
          if (promptLibraryAPI) promptLibraryAPI.refresh();
        } catch {
          showToast({ message: "Не удалось сохранить запрос", type: "error" });
        }
      }
    }

    if (msgAction === "continue_exec") {
      if (continueResolver) {
        continueResolver(true);
        continueResolver = null;
      }
    }

    if (msgAction === "stop_exec") {
      if (continueResolver) {
        continueResolver(false);
        continueResolver = null;
      }
    }
  });

  document.addEventListener("feedback", ((e: CustomEvent) => {
    const { feedback, messageId, message } = e.detail;
    const backendUrl = getBackendUrl();
    fetch(`${backendUrl}/api/feedback`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rating: feedback, messageId, comment: message }),
    }).catch(() => {
      /* silent */
    });
    if (feedback === "helpful") {
      showToast({ message: "Спасибо за отзыв!", type: "success" });
    } else {
      showToast({ message: "Спасибо! Мы станем лучше.", type: "info" });
    }
  }) as EventListener);

  // ── Keyboard shortcuts ──
  document.addEventListener("keydown", (e: KeyboardEvent) => {
    const target = e.target as HTMLElement;
    const isInput = target.tagName === "TEXTAREA" || target.tagName === "INPUT";

    // Ctrl+Enter / Cmd+Enter → отправить
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
      e.preventDefault();
      handleSend();
      return;
    }

    // Ctrl+K / Cmd+K → фокус на ввод
    if ((e.ctrlKey || e.metaKey) && e.key === "k") {
      e.preventDefault();
      const input = document.getElementById(
        "chat-input",
      ) as HTMLTextAreaElement;
      if (input) {
        input.focus();
      }
      return;
    }

    // Esc → остановить генерацию
    if (e.key === "Escape" && isProcessing) {
      if (reactLoop) {
        reactLoop.abort();
      }
      hideStopButton();
      showToast({ message: "Генерация остановлена", type: "warning" });
      return;
    }

    // Ctrl+Shift+C → copy selection info
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === "C") {
      e.preventDefault();
      updateSelectionDisplay();
      showToast({ message: "Информация о выделении обновлена", type: "info" });
    }

    // Ctrl+Z → Undo (not in text inputs)
    if ((e.ctrlKey || e.metaKey) && e.key === "z" && !isInput) {
      e.preventDefault();
      handleUndo();
    }
  });
}

async function testBackendConnection(url: string) {
  try {
    const response = await fetch(`${url}/api/health`, { mode: "cors" });
    if (response.ok) {
      updateConnectionStatus("online", "Подключено к бэкенду");
      if (headerAPI) headerAPI.setConnectionStatus("online");
    } else {
      updateConnectionStatus("offline", "Бэкенд нездоров");
      if (headerAPI) headerAPI.setConnectionStatus("offline");
    }
  } catch {
    updateConnectionStatus("offline", "Бэкенд недоступен");
    if (headerAPI) headerAPI.setConnectionStatus("offline");
    addMessage(
      "system",
      "⚠️ Не удаётся подключиться к серверу. Убедитесь, что он запущен (npm run dev в backend/).",
    );
    showToast({ message: "Бэкенд недоступен", type: "error" });
  }
}

// ── Narrow Mode Detection ──

function initNarrowMode() {
  const updateMode = () => {
    const w = window.innerWidth;
    document.body.classList.toggle("narrow-mode", w <= 420);
    document.body.classList.toggle("ultra-narrow-mode", w <= 360);
  };

  if (window.ResizeObserver) {
    const ro = new ResizeObserver(updateMode);
    ro.observe(document.body);
  }
  updateMode();
  window.addEventListener("resize", updateMode);
}

// ── Bootstrap ──

Office.onReady((info) => {
  try {
    if (info.host === Office.HostType.Excel) {
      // Нормализуем сохранённый backend_url ДО любого fetch.
      // В старых сборках в localStorage могло остаться http://localhost:4000 —
      // оно мгновенно блокируется как mixed content (Failed to fetch).
      const storedUrl = localStorage.getItem(BACKEND_URL_STORAGE_KEY);
      if (storedUrl && /^http:\/\/localhost/i.test(storedUrl)) {
        const fixed = storedUrl.replace(/^http:/i, "https:");
        localStorage.setItem(BACKEND_URL_STORAGE_KEY, fixed);
        console.info("[taskpane] normalized backend_url http→https:", fixed);
      }

      initTheme();
      initNarrowMode();
      initSession();

      updateConnectionStatus("loading", "Инициализация...");

      const app = document.getElementById("app");
      if (!app) throw new Error("App root not found");

      // ── Header ──

      headerAPI = createHeader({
        brandName: "AI Агент",
        model: DEFAULT_MODEL,
        connectionStatus: "loading",
        onSettingsClick: showSettingsDialog,
        onNewChat: handleNewChat,
        onModelChange: (model) => {
          localStorage.setItem(MODEL_STORAGE_KEY, model);
          if (llmClient) {
            llmClient.setModel(model);
          }
        },
        onExportChat: () => {
          const messagesDiv = document.getElementById("messages");
          if (!messagesDiv) return;
          const text = messagesDiv.textContent || "";
          navigator.clipboard.writeText(text).then(() => {
            addMessage("system", "✅ Чат экспортирован в буфер обмена");
            showToast({
              message: "Чат экспортирован в буфер",
              type: "success",
            });
          });
        },
        onFeedback: () => {
          showToast({
            message: "Отзыв: github.com/anomalyco/excel-ai-agent",
            type: "info",
          });
        },
        onUndo: handleUndo,
      });

      app.insertBefore(
        headerAPI.element,
        document.getElementById("connection-status"),
      );

      // ── Tab Bar ──

      tabBarAPI = createTabBar({
        tabs: TABS,
        activeTab: "chat",
        onTabChange: switchTab,
      });

      const tabBarEl = document.getElementById("tab-bar");
      if (tabBarEl) {
        tabBarEl.replaceWith(tabBarAPI.element);
      }

      // ── History Tab ──

      historyAPI = createHistory({
        onSelectSession: loadSession,
        onNewChat: handleNewChat,
      });

      const historyPanel = document.getElementById("panel-history");
      if (historyPanel && historyPanel.parentNode) {
        historyPanel.replaceWith(historyAPI.element);
        historyAPI.element.classList.remove("active");
      }

      // ── Prompt Library Tab ──

      promptLibraryAPI = createPromptLibrary({
        onSelectPrompt: (text: string, autoSend?: boolean) => {
          handlePromptSuggestion(text);
        },
      });

      const promptsPanel = document.getElementById("panel-prompts");
      if (promptsPanel && promptsPanel.parentNode) {
        promptsPanel.replaceWith(promptLibraryAPI.element);
        promptLibraryAPI.element.classList.remove("active");
      }

      // ── Input Area ──

      const inputAreaEl = document.getElementById("input-area");
      if (inputAreaEl) {
        inputAPI = createMessageInput({
          placeholder: "Спросите что-нибудь о таблице...",
          onSend: () => handleSend(),
          onChipClick: (text: string) => {
            handlePromptSuggestion(text);
          },
          onSelectionClick: () => {
            updateSelectionDisplay();
            const input = document.getElementById(
              "chat-input",
            ) as HTMLTextAreaElement;
            if (input) {
              const existingText = input.value.trim();
              const selection = existingText ? ` ${existingText}` : "";
              input.value = `Использовать выделенный диапазон${selection}`;
              input.dispatchEvent(new Event("input", { bubbles: true }));
              handleSend();
            }
          },
          onFileDrop: (files: File[]) => {
            addMessage(
              "system",
              `📎 Получено ${files.length} файл(ов): ${files.map((f) => f.name).join(", ")}`,
            );
            showToast({
              message: `📎 ${files.length} файл(ов) прикреплено`,
              type: "info",
            });
          },
        });
        inputAreaEl.replaceWith(inputAPI.element);
      }

      setupEventListeners();
      setupStopButton();
      setupModelSelector();
      setupSelectionButton();
      updateUndoButton();

      showWelcome();

      const backendUrl = getBackendUrl();
      if (backendUrl) {
        initLlmClient(backendUrl);
        initReActLoop();

        testBackendConnection(backendUrl);

        updateConnectionStatus("online", "Подключено");
        if (headerAPI) headerAPI.setConnectionStatus("online");

        addMessage(
          "system",
          "✅ Excel AI Агент готов! (Llama 3.3 70B на Groq)",
        );
        setInputEnabled(true);
      } else {
        updateConnectionStatus("offline", "Нет URL бэкенда");
        if (headerAPI) headerAPI.setConnectionStatus("offline");
        addMessage(
          "system",
          "Нажмите на шестерёнку ⚙, чтобы указать URL бэкенда.",
        );
        setInputEnabled(false);
      }
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    updateConnectionStatus("offline", "Ошибка");
    if (headerAPI) headerAPI.setConnectionStatus("offline");
    const el = document.createElement("div");
    el.style.cssText = "padding:16px;color:#d32f2f;font-size:12px";
    el.textContent = "Ошибка инициализации: " + msg;
    document.getElementById("messages")?.appendChild(el);
  }
});
