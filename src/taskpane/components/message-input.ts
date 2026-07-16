export interface SuggestedChip {
  icon: string;
  text: string;
}

export interface MessageInputOptions {
  placeholder?: string;
  onSend: (text: string) => void;
  suggestedChips?: SuggestedChip[];
  onChipClick?: (text: string) => void;
  onSelectionClick?: () => void;
  onFileDrop?: (files: File[]) => void;
}

export interface MessageInputAPI {
  element: HTMLElement;
  textarea: HTMLTextAreaElement;
  setEnabled(enabled: boolean): void;
  focus(): void;
  clear(): void;
  setSelectionLabel(label: string): void;
  setSelectionVisible(visible: boolean): void;
  setVoiceSupported(supported: boolean): void;
}

const DEFAULT_CHIPS: SuggestedChip[] = [
  { icon: "📊", text: "Analyze range" },
  { icon: "📈", text: "Create chart" },
  { icon: "🧹", text: "Clean data" },
  { icon: "📋", text: "Format as table" },
];

// ── Voice Input ──

const SpeechRecognitionAPI =
  window.SpeechRecognition || window.webkitSpeechRecognition;
const voiceSupported = !!SpeechRecognitionAPI;

let recognitionInstance: SpeechRecognition | null = null;
let currentLang = "auto";

const VOICE_LANGS = [
  { label: "RU", value: "ru-RU" },
  { label: "EN", value: "en-US" },
];

function getRecognitionLang(): string {
  if (currentLang === "auto") {
    const browserLang = navigator.language || "en-US";
    if (browserLang.startsWith("ru")) return "ru-RU";
    return "en-US";
  }
  return currentLang;
}

function cycleVoiceLang(): void {
  const idx = VOICE_LANGS.findIndex((l) => l.value === currentLang);
  currentLang = VOICE_LANGS[(idx + 1) % VOICE_LANGS.length].value;
}

function createVoiceButton(
  onTranscript: (text: string) => void,
  onStateChange: (recording: boolean) => void,
): { button: HTMLElement } {
  const btn = document.createElement("button");
  btn.className = "voice-btn";
  btn.title = "Voice input";
  btn.innerHTML =
    '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" y1="19" x2="12" y2="23"/><line x1="8" y1="23" x2="16" y2="23"/></svg>';
  btn.style.display = "none";

  btn.addEventListener("click", () => {
    if (!SpeechRecognitionAPI) return;

    if (recognitionInstance) {
      try {
        recognitionInstance.abort();
      } catch {}
      recognitionInstance = null;
      onStateChange(false);
      btn.classList.remove("recording");
      return;
    }

    const recognition = new SpeechRecognitionAPI();
    recognition.continuous = false;
    recognition.interimResults = true;
    recognition.lang = getRecognitionLang();

    recognition.onresult = (event: SpeechRecognitionEvent) => {
      let finalText = "";
      let interimText = "";
      for (let i = event.resultIndex; i < event.results.length; i++) {
        if (event.results[i].isFinal) {
          finalText += event.results[i][0].transcript;
        } else {
          interimText += event.results[i][0].transcript;
        }
      }
      const text = finalText || interimText;
      if (text) {
        onTranscript(text);
      }
    };

    recognition.onerror = () => {
      recognitionInstance = null;
      onStateChange(false);
      btn.classList.remove("recording");
    };

    recognition.onend = () => {
      recognitionInstance = null;
      onStateChange(false);
      btn.classList.remove("recording");
    };

    recognitionInstance = recognition;
    recognition.start();
    btn.classList.add("recording");
    onStateChange(true);
  });

  return { button: btn };
}

// ── Main Component ──

export function createMessageInput(
  options: MessageInputOptions,
): MessageInputAPI {
  const container = document.createElement("div");
  container.id = "input-area";
  container.classList.add("drop-target");

  const chips = options.suggestedChips || DEFAULT_CHIPS;

  let chipsHtml = "";
  if (chips.length > 0) {
    chipsHtml = '<div class="suggested-chips">';
    for (const chip of chips) {
      chipsHtml += `<button class="suggestion-chip" data-prompt="${escapeAttr(chip.text)}">
        <span class="chip-icon">${escapeHtml(chip.icon)}</span>
        <span>${escapeHtml(chip.text)}</span>
      </button>`;
    }
    chipsHtml += "</div>";
  }

  container.innerHTML = `
    <div class="input-area-top">
      <button class="selection-badge" id="selection-btn" style="display:none">
        📌 <span class="selection-label" id="selection-label">A1:B10</span>
      </button>
    </div>
    <div class="input-row">
      <textarea
        id="chat-input"
        rows="1"
        placeholder="${escapeAttr(options.placeholder || "Ask something about the sheet...")}"
        aria-label="Input field"
      ></textarea>
      <div class="input-actions">
        <div class="voice-group" id="voice-group" style="display:none">
          <button class="voice-lang-toggle" id="voice-lang-toggle" title="Recognition language">RU</button>
        </div>
        <button id="send-btn" class="send-btn" disabled aria-label="Send">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
        </button>
      </div>
    </div>
    ${chipsHtml}
    <div class="drop-indicator" id="drop-indicator">📎 Drop file here</div>
  `;

  const textarea = container.querySelector(
    "#chat-input",
  ) as HTMLTextAreaElement;
  const sendBtn = container.querySelector("#send-btn") as HTMLButtonElement;
  const selectionBtn = container.querySelector(
    "#selection-btn",
  ) as HTMLButtonElement;
  const dropIndicator = container.querySelector(
    "#drop-indicator",
  ) as HTMLElement;
  const voiceGroup = container.querySelector("#voice-group") as HTMLElement;
  const langToggle = container.querySelector(
    "#voice-lang-toggle",
  ) as HTMLButtonElement;

  // ── Voice Button ──

  const { button: voiceBtn } = createVoiceButton(
    (transcript: string) => {
      textarea.value += (textarea.value ? " " : "") + transcript;
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
    },
    (recording: boolean) => {
      container.classList.toggle("recording", recording);
    },
  );

  if (voiceSupported) {
    voiceGroup.style.display = "";
    voiceGroup.insertBefore(voiceBtn, langToggle);
  }

  if (langToggle) {
    langToggle.addEventListener("click", cycleVoiceLang);
  }

  // ── Drag & Drop ──

  let dragCounter = 0;

  container.addEventListener("dragenter", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter++;
    container.classList.add("dragging");
  });

  container.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.stopPropagation();
  });

  container.addEventListener("dragleave", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter--;
    if (dragCounter <= 0) {
      dragCounter = 0;
      container.classList.remove("dragging");
    }
  });

  container.addEventListener("drop", (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter = 0;
    container.classList.remove("dragging");

    if (e.dataTransfer?.files && e.dataTransfer.files.length > 0) {
      const files = Array.from(e.dataTransfer.files);
      options.onFileDrop?.(files);

      const fileNames = files.map((f) => f.name).join(", ");
      textarea.value +=
        (textarea.value ? " " : "") + `[Attached: ${fileNames}]`;
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
      showFileToast(`📎 ${files.length} file(s) attached`);
    }
  });

  function showFileToast(msg: string) {
    const existing = document.querySelector(".file-toast");
    if (existing) existing.remove();
    const toast = document.createElement("div");
    toast.className = "file-toast";
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 2500);
  }

  // ── Input handling ──

  function updateSendButton() {
    sendBtn.disabled = !textarea.value.trim();
  }

  textarea.addEventListener("input", () => {
    textarea.style.height = "auto";
    textarea.style.height = Math.min(textarea.scrollHeight, 100) + "px";
    updateSendButton();
  });

  textarea.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });

  sendBtn.addEventListener("click", handleSend);

  selectionBtn.addEventListener("click", () => {
    options.onSelectionClick?.();
  });

  function handleSend() {
    const text = textarea.value.trim();
    if (!text) return;
    options.onSend(text);
  }

  container.addEventListener("click", (e) => {
    const chip = (e.target as HTMLElement).closest(
      ".suggestion-chip",
    ) as HTMLElement;
    if (chip) {
      const prompt = chip.dataset.prompt;
      if (prompt) {
        options.onChipClick?.(prompt);
      }
    }
  });

  return {
    element: container,
    textarea,
    setEnabled(enabled: boolean) {
      textarea.disabled = !enabled;
      sendBtn.disabled = !enabled || !textarea.value.trim();
      if (enabled) {
        textarea.focus();
      }
    },
    focus() {
      textarea.focus();
    },
    clear() {
      textarea.value = "";
      textarea.style.height = "auto";
      updateSendButton();
    },
    setSelectionLabel(label: string) {
      const labelEl = container.querySelector(
        "#selection-label",
      ) as HTMLElement;
      if (labelEl) labelEl.textContent = label;
    },
    setSelectionVisible(visible: boolean) {
      selectionBtn.style.display = visible ? "" : "none";
    },
    setVoiceSupported(supported: boolean) {
      voiceGroup.style.display = supported ? "" : "none";
    },
  };
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
