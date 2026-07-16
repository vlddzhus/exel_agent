import { LlmClient, ChatMessage, ToolDefinition } from "../chat/llm-client";

// ── Structured Plan (JSON schema, replaces markdown plan) ──

export interface PlanStep {
  id: string;
  description: string;
  tool: string;
  args: Record<string, unknown>;
  dependsOn: string[];
  canRunInParallel: boolean;
  riskLevel: "safe" | "moderate" | "dangerous";
}

export interface StructuredPlan {
  title: string;
  summary: string;
  steps: PlanStep[];
  riskLevel: "low" | "medium" | "high";
  dataImpact: "read-only" | "modify" | "destructive";
  estimatedSteps: number;
}

// ── Execution Result ──

export interface StepResult {
  stepId: string;
  success: boolean;
  toolName: string;
  result: string;
  error?: string;
  elapsedMs: number;
}

export interface ExecutionResult {
  steps: StepResult[];
  totalElapsedMs: number;
  toolCallCount: number;
  rowsModified: number;
  toolNames: string[];
  errors: string[];
}

// ── Verification Result ──

export interface VerificationResult {
  verified: boolean;
  issues: VerificationIssue[];
  summary: string;
  canAutoFix: boolean;
}

export interface VerificationIssue {
  stepId: string;
  severity: "warning" | "error";
  description: string;
  suggestedFix?: string;
}

// ── Agent Context ──

export interface AgentContext {
  llm: LlmClient;
  messages: ChatMessage[];
  tools: ToolDefinition[];
  onMessage: MessageCallback;
  onThinking: ThinkingCallback;
  onProgress: ProgressCallback;
  onConfirm: ConfirmationHandler;
  aborted: boolean;
}

export type MessageCallback = (
  role: string,
  content: string,
  actionButtons?: { label: string; id: string }[],
) => void;

export type ThinkingCallback = (
  text: string,
  phase?: "reasoning" | "executing" | "verifying" | "summarizing",
) => void;

export type ProgressCallback = (
  current: number,
  total: number,
  label: string,
) => void;

export type ConfirmationHandler = (
  toolName: string,
  args: Record<string, unknown>,
  description?: string,
) => Promise<boolean>;

// ── Agent Interface ──

export interface AgentResult<T = unknown> {
  success: boolean;
  data: T;
  error?: string;
  tokensUsed: number;
}

export interface Agent {
  readonly name: string;
  readonly model: string;
  execute(input: AgentInput, context: AgentContext): Promise<AgentResult>;
}

export interface AgentInput {
  userMessage: string;
  conversationHistory: ChatMessage[];
  plan?: StructuredPlan;
  executionResult?: ExecutionResult;
  verificationResult?: VerificationResult;
}

// ── Model Configuration ──

export interface ModelConfig {
  supervisor: string;
  planner: string;
  executor: string;
  verifier: string;
  summarizer: string;
}

export const DEFAULT_MODEL_CONFIG: ModelConfig = {
  supervisor: "gpt-4o-mini",
  planner: "gpt-4o",
  executor: "gpt-4o",
  verifier: "gpt-4o-mini",
  summarizer: "gpt-4o-mini",
};
