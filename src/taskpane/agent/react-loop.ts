import { LlmClient, ChatMessage } from "../chat/llm-client";
import { Orchestrator, OrchestratorConfig } from "../agents/orchestrator";
import { StructuredPlan } from "../agents/types";
import { PlanStep } from "../components/plan-card";

// ── Public API (backward-compatible with taskpane.ts) ──

export type ConfirmationHandler = (
  toolName: string,
  args: Record<string, unknown>,
  description?: string,
) => Promise<boolean>;
export type ContinueHandler = () => Promise<boolean>;

export type MessageCallback = (
  role: string,
  content: string,
  actionButtons?: { label: string; id: string }[],
) => void;

export type PlanConfirmHandler = (
  planText: string,
  steps: PlanStep[],
) => Promise<{ approved: boolean; editedPlan?: string }>;

export type ThinkingCallback = (
  text: string,
  phase?: "reasoning" | "executing" | "verifying" | "summarizing",
) => void;

export type AbortCallback = () => void;

export interface ExecutionStats {
  toolCallCount: number;
  elapsedMs: number;
  rowsModified: number;
  toolNames: string[];
  finalMessage: string;
  totalTokens: number;
}

export type ExecutionCompleteCallback = (stats: ExecutionStats) => void;

// ── Intent Detection (delegates to Supervisor agent) ──

export async function detectIntent(
  userInput: string,
  llm?: LlmClient,
): Promise<"chat" | "task"> {
  const text = userInput.trim();
  const wordCount = text.split(/\s+/).length;

  // Fast path: obvious short greetings
  const QUICK_GREETINGS = [
    /^(привет|здравствуй|здравствуйте|хай|хелло|дарова|салют|ку|здарова)\b/i,
    /^(hello|hi|hey|greetings|good morning|good afternoon|good evening)\b/i,
    /^(спасибо|благодарю|благодарствую|thanks|thank you)\b/i,
    /^(пока|до свидания|увидимся|bye|goodbye)\b/i,
    /^(как дела|как ты|how are you|what'?s up|how'?s it going)\b/i,
    /^(who are you|расскажи о себе|что ты умеешь)\b/i,
  ];

  if (wordCount <= 3) {
    for (const p of QUICK_GREETINGS) {
      if (p.test(text)) return "chat";
    }
  }

  // LLM-based classification
  if (llm) {
    try {
      const response = await llm.chat([
        {
          role: "system",
          content:
            'You are an intent classifier for an Excel AI agent. Reply with ONE word only.\n- "task": user wants to DO something in Excel (modify, analyze, format, clean, calculate, create, etc.)\n- "chat": greeting, thanks, small talk, or question about the agent',
        },
        { role: "user", content: text },
      ]);
      const result =
        response.choices[0]?.message?.content?.trim().toLowerCase() || "";
      if (result === "task") return "task";
      if (result === "chat") return "chat";
    } catch {
      // fall through to regex
    }
  }

  // Fallback: regex
  const TASK_VERBS =
    /(сделай|сделат|напиш|напис|посчитай|посчитат|покаж|показа|найди|найти|удали|удал|добав|измен|создай|созда|расшир|переимен|скопируй|встав|анализ|проанализ|отсортир|сортир|отфильтр|фильтр|построй|нарисуй|диаграмм|график|сумм|суммиру|таблиц|сводн|формат|отчет|отчёт|очист|дубликат)/i;
  const TASK_ENGLISH =
    /\b(clean|sort|filter|chart|graph|table|format|analyze|calculate|sum|total|average|find|show|create|make|add|remove|update|set|insert|delete|copy|paste|move|export|import|pivot|vlookup|formula)\b/i;
  const EXCEL_KEYWORDS =
    /\b(excel|таблиц|лист|ячейк|данн|range|cell|column|row|sheet|worksheet|workbook|data|values)\b/i;

  if (TASK_VERBS.test(text)) return "task";
  if (TASK_ENGLISH.test(text)) return "task";
  if (EXCEL_KEYWORDS.test(text)) return "task";

  if (wordCount <= 5) {
    if (
      /^(что ты|как ты|зачем|почему|когда|где|кто|какой|какая|какие)\b/i.test(
        text,
      )
    )
      return "chat";
    if (/^(а что|а как|а зачем)\b/i.test(text)) return "chat";
    if (/^(расскажи|объясни|объясните)\b/i.test(text)) return "chat";
  }

  if (wordCount <= 2) return "chat";

  return "task";
}

// ── ReActLoop — now a thin wrapper around Orchestrator ──

export class ReActLoop {
  private orchestrator: Orchestrator;
  private onMessage: MessageCallback;
  private onComplete: ExecutionCompleteCallback | null;
  private planConfirmHandler: PlanConfirmHandler;

  constructor(
    llm: LlmClient,
    onMessage: MessageCallback,
    onConfirm: ConfirmationHandler,
    onPlanConfirm: PlanConfirmHandler,
    onProgress?: (current: number, total: number, label: string) => void,
    onThinking?: ThinkingCallback,
    onComplete?: ExecutionCompleteCallback,
    _onContinue?: ContinueHandler,
  ) {
    this.onMessage = onMessage;
    this.onComplete = onComplete ?? null;
    this.planConfirmHandler = onPlanConfirm;

    const config: OrchestratorConfig = {
      llm,
      onMessage,
      onThinking: onThinking ?? (() => {}),
      onProgress: onProgress ?? (() => {}),
      onConfirm,
      onPlanConfirm: async (plan: StructuredPlan) => {
        const planText = Orchestrator.planToText(plan);
        const steps = Orchestrator.planToSteps(plan);
        const result = await onPlanConfirm(planText, steps);
        return {
          approved: result.approved,
          editedPlan: result.editedPlan ? undefined : undefined, // editedPlan is text-based, keep plan as-is
        };
      },
      onComplete: (stats) => {
        if (this.onComplete) {
          this.onComplete({
            toolCallCount: stats.toolCallCount,
            elapsedMs: stats.elapsedMs,
            rowsModified: stats.rowsModified,
            toolNames: stats.toolNames,
            finalMessage: stats.finalMessage,
            totalTokens: stats.totalTokens,
          });
        }
      },
    };

    this.orchestrator = new Orchestrator(config);
  }

  abort() {
    this.orchestrator.abort();
  }

  async handleChat(userInput: string): Promise<string> {
    return this.orchestrator.handleChat(userInput);
  }

  async planAndExecute(userInput: string): Promise<string> {
    return this.orchestrator.handleTask(userInput);
  }

  // Legacy alias
  async execute(userInput: string): Promise<string> {
    return this.orchestrator.handleTask(userInput);
  }

  clearConversation() {
    this.orchestrator.clearConversation();
  }

  getMessages(): ChatMessage[] {
    return this.orchestrator.getMessages();
  }
}
