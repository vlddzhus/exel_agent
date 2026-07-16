export interface HeaderOptions {
  brandName: string;
  model: string;
  connectionStatus: "online" | "offline" | "loading";
  onSettingsClick: () => void;
  onNewChat: () => void;
  onModelChange?: (model: string) => void;
  onExportChat?: () => void;
  onFeedback?: () => void;
  onUndo?: () => void;
}

export interface HeaderAPI {
  element: HTMLElement;
  setModel(model: string): void;
  setConnectionStatus(status: "online" | "offline" | "loading"): void;
  setUndoCount(count: number): void;
}

const MODEL_OPTIONS: { value: string; label: string }[] = [
  { value: "openai/gpt-5-4", label: "GPT-5.4" },
  { value: "openai/gpt-5-4-pro", label: "GPT-5.4 Pro" },
];

export function createHeader(options: HeaderOptions): HeaderAPI {
  const header = document.createElement("header");
  header.id = "header";

  header.innerHTML = `
    <div class="header-left">
      <span class="brand-dot"></span>
      <span class="brand-name">${escapeHtml(options.brandName)}</span>
    </div>
    <div class="header-center">
      <select id="model-select" class="model-select" aria-label="Select model">
        ${MODEL_OPTIONS.map((m) => `<option value="${m.value}" ${m.value === options.model ? "selected" : ""}>${escapeHtml(m.label)}</option>`).join("")}
      </select>
      <span class="connection-indicator ${options.connectionStatus}" id="header-connection-dot" title="Connection status"></span>
    </div>
    <div class="header-right">
      <button id="new-chat-btn" class="header-btn" title="New chat" aria-label="New chat">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
      </button>
      <button id="settings-btn" class="header-btn" title="Settings" aria-label="Settings">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/></svg>
      </button>
      <button id="overflow-btn" class="header-btn" title="More" aria-label="More">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="1"/><circle cx="19" cy="12" r="1"/><circle cx="5" cy="12" r="1"/></svg>
      </button>
    </div>
  `;

  let overflowMenu: HTMLElement | null = null;

  const newChatBtn = header.querySelector("#new-chat-btn") as HTMLButtonElement;
  newChatBtn.addEventListener("click", () => options.onNewChat());

  const overflowBtn = header.querySelector(
    "#overflow-btn",
  ) as HTMLButtonElement;
  overflowBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleOverflowMenu();
  });

  function toggleOverflowMenu() {
    if (overflowMenu) {
      closeOverflowMenu();
      return;
    }
    overflowMenu = document.createElement("div");
    overflowMenu.className = "overflow-menu";
    overflowMenu.innerHTML = `
      <button class="overflow-item" data-action="export"><span class="icon">↓</span> Export chat</button>
      <button class="overflow-item" data-action="undo" id="overflow-undo"><span class="icon">↩</span> Undo</button>
      <div class="overflow-divider"></div>
      <button class="overflow-item" data-action="feedback"><span class="icon">💬</span> Send feedback</button>
    `;
    document.body.appendChild(overflowMenu);

    overflowMenu.addEventListener("click", (e) => {
      const target = e.target as HTMLElement;
      const action = target.dataset.action;
      if (!action) return;
      closeOverflowMenu();
      switch (action) {
        case "export":
          options.onExportChat?.();
          break;
        case "undo":
          options.onUndo?.();
          break;
        case "feedback":
          options.onFeedback?.();
          break;
      }
    });

    setTimeout(() => {
      document.addEventListener("click", closeOverflowMenu, { once: true });
    }, 0);
  }

  function closeOverflowMenu() {
    if (overflowMenu) {
      overflowMenu.remove();
      overflowMenu = null;
    }
  }

  const modelSelect = header.querySelector(
    "#model-select",
  ) as HTMLSelectElement;
  modelSelect.addEventListener("change", () => {
    options.onModelChange?.(modelSelect.value);
  });

  const settingsBtn = header.querySelector(
    "#settings-btn",
  ) as HTMLButtonElement;
  settingsBtn.addEventListener("click", () => options.onSettingsClick());

  return {
    element: header,
    setModel(model: string) {
      modelSelect.value = model;
    },
    setConnectionStatus(status: "online" | "offline" | "loading") {
      const dot = header.querySelector("#header-connection-dot") as HTMLElement;
      if (dot) {
        dot.className = `connection-indicator ${status}`;
      }
    },
    setUndoCount(count: number) {
      const undoItem = document.getElementById("overflow-undo");
      if (undoItem) {
        undoItem.style.display = count > 0 ? "" : "none";
      }
    },
  };
}

function escapeHtml(s: string): string {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}
