const DEFAULT_BACKEND_URL = "https://localhost:4000";
const DEFAULT_MODEL = "llama-3.3-70b-versatile";

export const MODEL_CHAIN = [
  "llama-3.3-70b-versatile",
  "llama-4-scout-17b-16e-instruct",
  "qwen-qwen3-32b",
  "deepseek-r1-distill-70b",
];

export interface LlmConfig {
  backendUrl: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
}

export interface ChatMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string;
  tool_call_id?: string;
  tool_calls?: ToolCall[];
}

export interface ToolCall {
  id: string;
  type: "function";
  function: {
    name: string;
    arguments: string;
  };
}

export interface ToolDefinition {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface LlmResponse {
  id?: string;
  choices: {
    index: number;
    message: {
      role: string;
      content: string | null;
      tool_calls?: ToolCall[];
    };
    finish_reason: string;
  }[];
  usage?: {
    total_tokens?: number;
    input_tokens?: number;
    output_tokens?: number;
    cost?: string;
  };
}

type ModelFallbackCallback = (oldModel: string, newModel: string) => void;

export class LlmClient {
  private config: LlmConfig;
  private modelIndex: number;
  private onModelFallback: ModelFallbackCallback | null = null;
  private fallbackAttempts = 0;
  private maxFallbackAttempts = 3;
  private abortController: AbortController | null = null;
  private currentStreamReader: ReadableStreamDefaultReader<Uint8Array> | null =
    null;

  constructor(config: LlmConfig) {
    this.config = config;
    const idx = MODEL_CHAIN.indexOf(this.config.model);
    this.modelIndex = idx >= 0 ? idx : 0;
    this.config.model = MODEL_CHAIN[this.modelIndex];
  }

  setBackendUrl(url: string) {
    this.config.backendUrl = url;
  }

  setModel(model: string) {
    this.config.model = model;
    const idx = MODEL_CHAIN.indexOf(model);
    this.modelIndex = idx >= 0 ? idx : 0;
    this.fallbackAttempts = 0;
  }

  getModel(): string {
    return this.config.model || DEFAULT_MODEL;
  }
  setOnModelFallback(callback: ModelFallbackCallback) {
    this.onModelFallback = callback;
  }
  getModelChain(): string[] {
    return MODEL_CHAIN;
  }
  getCurrentModelIndex(): number {
    return this.modelIndex;
  }
  isConfigured(): boolean {
    return !!this.config.backendUrl;
  }
  getBackendUrl(): string {
    return this.config.backendUrl || DEFAULT_BACKEND_URL;
  }
  getFallbackAttempts(): number {
    return this.fallbackAttempts;
  }
  resetFallbackAttempts() {
    this.fallbackAttempts = 0;
  }

  abort() {
    if (this.currentStreamReader) {
      try {
        this.currentStreamReader.cancel();
      } catch {}
      this.currentStreamReader = null;
    }
    if (this.abortController) {
      this.abortController.abort();
      this.abortController = null;
    }
  }

  private async fetchWithRetry(
    url: string,
    options: RequestInit,
    retries = 3,
  ): Promise<Response> {
    this.abortController = new AbortController();
    const fetchOptions = { ...options, signal: this.abortController.signal };

    for (let attempt = 0; attempt < retries; attempt++) {
      try {
        const response = await fetch(url, fetchOptions);
        if (response.ok) return response;

        if (response.status === 429 || response.status === 503) {
          if (attempt < retries - 1) {
            await new Promise((r) =>
              setTimeout(r, Math.pow(2, attempt + 1) * 1000),
            );
            continue;
          }
        }

        const errorText = await response.text();
        throw new Error(`Backend error (${response.status}): ${errorText}`);
      } catch (error: unknown) {
        if (error instanceof DOMException && error.name === "AbortError") {
          throw new Error("Generation cancelled by user.");
        }
        if (error instanceof TypeError) {
          throw new Error(
            "Не удаётся подключиться к серверу.\n" +
              "Проверьте:\n" +
              "1. Запущен ли backend (npm run dev в backend/)\n" +
              "2. Правильный ли URL бэкенда (шестерёнка в правом верхнем углу)\n" +
              "3. Работает ли интернет-соединение",
          );
        }
        if (attempt < retries - 1) {
          await new Promise((r) =>
            setTimeout(r, Math.pow(2, attempt + 1) * 1000),
          );
          continue;
        }
        throw error;
      }
    }
    throw new Error("Бэкенд: превышено количество попыток");
  }

  private tryFallbackModel(): boolean {
    if (this.modelIndex < MODEL_CHAIN.length - 1) {
      const oldModel = this.config.model;
      this.modelIndex++;
      this.config.model = MODEL_CHAIN[this.modelIndex];
      this.fallbackAttempts++;
      this.onModelFallback?.(oldModel, this.config.model);
      return true;
    }
    return false;
  }

  // ── Non-streaming chat ──

  async chat(
    messages: ChatMessage[],
    tools?: ToolDefinition[],
  ): Promise<LlmResponse> {
    const body: Record<string, unknown> = {
      model: this.config.model,
      messages,
      temperature: this.config.temperature ?? 0.3,
      max_tokens: this.config.maxTokens ?? 16384,
    };
    if (tools && tools.length > 0) body.tools = tools;

    const response = await this.fetchWithRetry(
      `${this.getBackendUrl()}/api/agent/complete`,
      {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      },
    );
    return response.json();
  }

  // ── SSE Streaming ──

  private async fetchStream(
    url: string,
    options: RequestInit,
    retries = 2,
  ): Promise<Response> {
    for (let attempt = 0; attempt < retries; attempt++) {
      try {
        const response = await fetch(url, options);
        if (response.ok) return response;
        if (response.status >= 400 && response.status < 500) {
          throw new Error(
            `Backend error (${response.status}): ${await response.text()}`,
          );
        }
        if (attempt < retries - 1) {
          await new Promise((r) =>
            setTimeout(r, Math.pow(2, attempt + 1) * 500),
          );
        }
      } catch (error: unknown) {
        if (error instanceof DOMException && error.name === "AbortError") {
          throw new Error("Generation cancelled by user.");
        }
        if (attempt < retries - 1) {
          await new Promise((r) =>
            setTimeout(r, Math.pow(2, attempt + 1) * 500),
          );
          continue;
        }
        throw error;
      }
    }
    throw new Error("Backend stream: exceeded retry attempts");
  }

  async chatStream(
    messages: ChatMessage[],
    tools: ToolDefinition[] | undefined,
    onChunk: (text: string) => void,
  ): Promise<LlmResponse> {
    const body: Record<string, unknown> = {
      model: this.config.model,
      messages,
      temperature: this.config.temperature ?? 0.3,
      max_tokens: this.config.maxTokens ?? 16384,
    };
    if (tools && tools.length > 0) body.tools = tools;

    this.abortController = new AbortController();
    const signal = this.abortController.signal;

    const response = await this.fetchStream(
      `${this.getBackendUrl()}/api/agent/stream`,
      {
        method: "POST",
        mode: "cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal,
      },
    );

    const responseBody = response.body;
    if (!responseBody) {
      throw new Error("Потоковая передача не поддерживается браузером");
    }

    const reader = responseBody.getReader();
    const self = this;
    self.currentStreamReader = reader;
    const decoder = new TextDecoder();
    let buffer = "";
    let pendingEventType = "";

    return new Promise<LlmResponse>((resolve, reject) => {
      const abortHandler = () => {
        reader.cancel().catch(() => {});
        self.currentStreamReader = null;
        reject(new Error("Generation cancelled by user."));
      };
      signal.addEventListener("abort", abortHandler, { once: true });

      function processEvent(eventType: string, dataStr: string) {
        try {
          const data = JSON.parse(dataStr);

          switch (eventType) {
            case "thinking":
              if (data.text) onChunk(data.text);
              break;
            case "status":
              if (data.phase === "fallback") {
                onChunk(
                  `\n\n*⚠️ Switching provider: ${data.from} → ${data.to}*`,
                );
              }
              break;
            case "done":
              reader.cancel().catch(() => {});
              self.currentStreamReader = null;
              resolve(data as LlmResponse);
              break;
            case "error":
              reject(new Error(data.error || "Stream error"));
              break;
          }
        } catch {
          // skip malformed JSON
        }
      }

      function processBuffer() {
        const endsWithNewline = buffer.endsWith("\n");
        const lines = buffer.split(/\r?\n/);

        if (!endsWithNewline) {
          buffer = lines.pop() ?? "";
        } else {
          buffer = "";
          lines.pop();
        }

        let eventType = pendingEventType;
        pendingEventType = "";
        let dataStr = "";

        for (const line of lines) {
          const ln = line.trim();
          if (!ln) {
            if (dataStr) {
              processEvent(eventType, dataStr);
              dataStr = "";
              eventType = "";
            }
            continue;
          }
          if (ln.startsWith("event: ")) {
            if (dataStr) {
              processEvent(eventType, dataStr);
              dataStr = "";
            }
            eventType = ln.slice(7);
          } else if (ln.startsWith("data: ")) {
            if (dataStr) dataStr += "\n";
            dataStr += ln.slice(6);
          }
        }

        if (dataStr) {
          processEvent(eventType, dataStr);
          eventType = "";
        }

        pendingEventType = eventType;
      }

      function pump(): void {
        reader
          .read()
          .then(({ done, value }) => {
            if (done) {
              reject(new Error("Stream ended unexpectedly"));
              return;
            }

            buffer += decoder.decode(value, { stream: true });
            processBuffer();
            pump();
          })
          .catch((err: Error) => {
            if (err.name === "AbortError") {
              reject(new Error("Generation cancelled by user."));
            } else {
              reject(err);
            }
          });
      }

      pump();
    }).finally(() => {
      self.currentStreamReader = null;
    });
  }
}
