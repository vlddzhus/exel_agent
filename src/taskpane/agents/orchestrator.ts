import { LlmClient, ChatMessage } from "../chat/llm-client";
import { toolRegistry } from "../tools/registry";
import { undoManager } from "../tools/backup";
import { CHAT_PROMPT } from "../agent/system-prompt";
import {
  Agent,
  AgentInput,
  AgentContext,
  StructuredPlan,
  ExecutionResult,
  VerificationResult,
  ModelConfig,
  DEFAULT_MODEL_CONFIG,
  MessageCallback,
  ThinkingCallback,
  ProgressCallback,
  ConfirmationHandler,
} from "./types";
import {
  SupervisorAgent,
  PlannerAgent,
  ExecutorAgent,
  VerifierAgent,
  SummarizerAgent,
} from "./agents";
import { PlanStep } from "../components/plan-card";

export interface OrchestratorConfig {
  llm: LlmClient;
  models?: Partial<ModelConfig>;
  onMessage: MessageCallback;
  onThinking: ThinkingCallback;
  onProgress: ProgressCallback;
  onConfirm: ConfirmationHandler;
  onPlanConfirm: (
    plan: StructuredPlan,
  ) => Promise<{ approved: boolean; editedPlan?: StructuredPlan }>;
  onComplete?: (stats: OrchestratorStats) => void;
  onContinue?: () => Promise<boolean>;
}

export interface OrchestratorStats {
  intent: "chat" | "task";
  toolCallCount: number;
  elapsedMs: number;
  rowsModified: number;
  toolNames: string[];
  finalMessage: string;
  totalTokens: number;
  verified: boolean;
  agentsUsed: string[];
}

export class Orchestrator {
  private supervisor: SupervisorAgent;
  private planner: PlannerAgent;
  private executor: ExecutorAgent;
  private verifier: VerifierAgent;
  private summarizer: SummarizerAgent;

  private llm: LlmClient;
  private messages: ChatMessage[] = [];
  private aborted = false;
  private config: OrchestratorConfig;
  private models: ModelConfig;

  // Stats
  private startTime = 0;
  private totalTokens = 0;
  private agentsUsed: string[] = [];

  constructor(config: OrchestratorConfig) {
    this.config = config;
    this.llm = config.llm;
    this.models = { ...DEFAULT_MODEL_CONFIG, ...config.models };

    this.supervisor = new SupervisorAgent(this.models.supervisor);
    this.planner = new PlannerAgent(this.models.planner);
    this.executor = new ExecutorAgent(this.models.executor);
    this.verifier = new VerifierAgent(this.models.verifier);
    this.summarizer = new SummarizerAgent(this.models.summarizer);

    this.messages = [{ role: "system", content: CHAT_PROMPT }];
  }

  abort() {
    this.aborted = true;
    this.llm.abort();
  }

  clearConversation() {
    this.messages = [{ role: "system", content: CHAT_PROMPT }];
    undoManager.clear();
    this.aborted = false;
    this.totalTokens = 0;
    this.agentsUsed = [];
  }

  getMessages(): ChatMessage[] {
    return this.messages;
  }

  private createContext(): AgentContext {
    return {
      llm: this.llm,
      messages: this.messages,
      tools: toolRegistry.getSchemas(),
      onMessage: this.config.onMessage,
      onThinking: this.config.onThinking,
      onProgress: this.config.onProgress,
      onConfirm: this.config.onConfirm,
      aborted: this.aborted,
    };
  }

  private buildAgentInput(userMessage: string): AgentInput {
    return {
      userMessage,
      conversationHistory: this.messages,
    };
  }

  // ── Chat mode (simple conversation) ──

  async handleChat(userInput: string): Promise<string> {
    this.aborted = false;
    this.startTime = Date.now();
    this.agentsUsed = ["supervisor"];

    this.messages.push({ role: "user", content: userInput });
    this.config.onThinking?.("🤔 Думаю...", "reasoning");

    try {
      const response = await this.llm.chat(this.messages, []);
      const content = response.choices[0]?.message?.content ?? "...";

      if (response.usage?.total_tokens) {
        this.totalTokens += response.usage.total_tokens;
      }

      this.messages.push({ role: "assistant", content });
      this.config.onMessage("assistant", content);

      this.fireComplete("chat", content, false);
      return content;
    } catch (error) {
      const errMsg = error instanceof Error ? error.message : String(error);
      if (errMsg.includes("cancelled") || errMsg.includes("abort")) {
        this.config.onMessage("system", "Generation cancelled.");
        this.fireComplete("chat", "Cancelled", false);
        return "Cancelled";
      }
      throw error;
    }
  }

  // ── Task mode (full multi-agent pipeline) ──

  async handleTask(userInput: string): Promise<string> {
    this.aborted = false;
    this.startTime = Date.now();
    this.agentsUsed = [];

    const context = this.createContext();
    const input = this.buildAgentInput(userInput);

    // ── Phase 1: Supervisor — classify intent ──
    this.config.onThinking?.("🧭 Фаза 1: Анализ запроса...", "reasoning");
    this.config.onProgress?.(1, 5, "Анализ запроса...");

    const supervisorResult = await this.supervisor.execute(input, context);
    this.agentsUsed.push("supervisor");
    this.totalTokens += supervisorResult.tokensUsed;

    if (supervisorResult.data === "chat") {
      return this.handleChat(userInput);
    }

    // ── Phase 2: Planner — create structured plan ──
    this.config.onThinking?.("📋 Фаза 2: Составление плана...", "reasoning");
    this.config.onProgress?.(2, 5, "Составление плана...");

    const plannerResult = await this.planner.execute(input, context);
    this.agentsUsed.push("planner");
    this.totalTokens += plannerResult.tokensUsed;

    const plan = plannerResult.data;
    if (!plannerResult.success) {
      this.config.onMessage(
        "system",
        `⚠️ Планирование не удалось: ${plannerResult.error}. Прямое выполнение.`,
      );
    }

    // ── Phase 3: Executor — execute the plan (no user approval, undo is safety) ──
    this.config.onThinking?.("🚀 Выполнение...", "executing");
    this.config.onProgress?.(3, 5, "Выполнение операций...");

    const execInput: AgentInput = {
      ...input,
      plan,
    };

    const execResult = await this.executor.execute(execInput, context);
    this.agentsUsed.push("executor");
    this.totalTokens += execResult.tokensUsed;

    const execution: ExecutionResult = execResult.data;

    // ── Phase 4: Verifier — check results ──
    this.config.onThinking?.("🔍 Фаза 4: Проверка результатов...", "verifying");
    this.config.onProgress?.(4, 5, "Проверка результатов...");

    const verifyInput: AgentInput = {
      ...input,
      plan,
      executionResult: execution,
    };

    const verifyResult = await this.verifier.execute(verifyInput, context);
    this.agentsUsed.push("verifier");
    this.totalTokens += verifyResult.tokensUsed;

    const verification: VerificationResult = verifyResult.data;

    // ── Phase 5: Summarizer — final report ──
    this.config.onThinking?.("📝 Фаза 5: Подготовка отчёта...", "summarizing");
    this.config.onProgress?.(5, 5, "Подготовка отчёта...");

    const summaryInput: AgentInput = {
      ...input,
      plan,
      executionResult: execution,
      verificationResult: verification,
    };

    const summaryResult = await this.summarizer.execute(summaryInput, context);
    this.agentsUsed.push("summarizer");
    this.totalTokens += summaryResult.tokensUsed;

    const finalMessage = summaryResult.data;
    this.messages.push({ role: "assistant", content: finalMessage });
    this.config.onMessage("assistant", finalMessage);

    this.fireComplete("task", finalMessage, verification.verified);

    return finalMessage;
  }

  private fireComplete(
    intent: "chat" | "task",
    finalMessage: string,
    verified: boolean,
  ) {
    if (this.config.onComplete) {
      this.config.onComplete({
        intent,
        toolCallCount: 0,
        elapsedMs: Date.now() - this.startTime,
        rowsModified: 0,
        toolNames: [],
        finalMessage,
        totalTokens: this.totalTokens,
        verified,
        agentsUsed: [...this.agentsUsed],
      });
    }
  }

  // ── Convert StructuredPlan to PlanStep[] for UI compatibility ──

  static planToSteps(plan: StructuredPlan): PlanStep[] {
    return plan.steps.map((s, i) => ({
      number: i + 1,
      summary: s.description,
      details: `Tool: ${s.tool}\nRisk: ${s.riskLevel}\nParallel: ${s.canRunInParallel ? "yes" : "no"}`,
      icon: assignIcon(s.description),
      status: "pending" as const,
    }));
  }

  static planToText(plan: StructuredPlan): string {
    const lines = plan.steps.map(
      (s, i) => `${i + 1}. ${s.description} [${s.tool}]`,
    );
    return `${plan.title}\n\n${lines.join("\n")}`;
  }
}

function assignIcon(text: string): string {
  const lower = text.toLowerCase();
  if (/\b(chart|graph|plot|visualize|dashboard)\b/.test(lower)) return "📈";
  if (/\b(analy|summar|statistic|mean|avg|total|count)\b/.test(lower))
    return "📊";
  if (/\b(clean|duplicate|remove|delete|empty|null|missing)\b/.test(lower))
    return "🧹";
  if (/\b(format|style|color|bold|header|border|align)\b/.test(lower))
    return "🎨";
  if (/\b(sort|filter|order|group)\b/.test(lower)) return "📋";
  if (/\b(report|export|save|output|print)\b/.test(lower)) return "📝";
  if (/\b(formula|calculate|compute|sum|if|vlookup)\b/.test(lower)) return "🔢";
  if (/\b(pivot|table|crosstab|aggregate)\b/.test(lower)) return "📊";
  if (/\b(read|load|import|fetch|get|extract)\b/.test(lower)) return "📂";
  if (/\b(update|set|write|put|change|modify|edit)\b/.test(lower)) return "✏️";
  if (/\b(check|verify|validate|test|review|confirm)\b/.test(lower))
    return "🔍";
  return "🔧";
}
