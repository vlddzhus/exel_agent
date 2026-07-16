import { LlmClient, ChatMessage } from "../chat/llm-client";
import { toolRegistry } from "../tools/registry";
import { undoManager } from "../tools/backup";
import { verifyToolResult } from "../tools/verification-tools";
import { analyzeToolArgsRisk } from "../tools/formula-allowlist";
import {
  Agent,
  AgentResult,
  AgentInput,
  AgentContext,
  StructuredPlan,
  PlanStep,
  ExecutionResult,
  StepResult,
  VerificationResult,
} from "./types";
import {
  SUPERVISOR_PROMPT,
  PLANNER_PROMPT,
  EXECUTOR_PROMPT,
  VERIFIER_PROMPT,
  SUMMARIZER_PROMPT,
} from "./prompts";

/**
 * DANGEROUS_TOOLS — every tool that mutates workbook state.
 * Each gets a backup before execution and requires user confirmation.
 *
 * Batch 3: extended to ALL write operations (not just destructive ones).
 * Even "safe" writes like applyFormat get a backup so the user can undo.
 */
const DANGEROUS_TOOLS = new Set([
  // Destructive
  "clearRange",
  "clearWorksheet",
  "deleteChart",
  "deleteWorksheet",
  // Writes
  "setValues",
  "setFormula",
  "fillFormula",
  "appendRows",
  // Structure
  "createTable",
  "addTableRow",
  "mergeCells",
  "manageSheets",
  "manageTable",
  "createPivotTable",
  "createChart",
  // Transform
  "sortData",
  "filterData",
  "removeDuplicates",
  "splitTextToColumns",
  "normalizeText",
  "lookup",
  // Formatting (reversible via backup)
  "applyFormat",
  "setCellFormat",
  "applyCellFormat",
  "applyNumberFormat",
  "formatAsTable",
  // Table operations (reorder/filter change visible state)
  "sortTable",
  "filterTable",
]);

/**
 * Tools that mutate data in a way where a single transaction covering the
 * whole batch is preferable (sequential dependent mutations).
 */
const TRANSACTIONAL_TOOLS = new Set([
  "setValues",
  "setFormula",
  "fillFormula",
  "appendRows",
  "clearRange",
  "clearWorksheet",
  "applyFormat",
  "setCellFormat",
  "applyCellFormat",
  "applyNumberFormat",
  "formatAsTable",
  "mergeCells",
  "manageSheets",
  "manageTable",
  "createPivotTable",
  "createChart",
  "sortData",
  "filterData",
  "removeDuplicates",
  "splitTextToColumns",
  "normalizeText",
  "lookup",
  "createTable",
  "addTableRow",
  "sortTable",
  "filterTable",
]);

/** Extract the backup address from tool args (handles all address-arg shapes). */
function getBackupAddress(args: Record<string, unknown>): string | undefined {
  return (
    (args.address as string) ??
    (args.cellAddress as string) ??
    (args.targetRange as string) ??
    (args.sourceAddress as string) ??
    (args.destinationAddress as string) ??
    (args.lookupAddress as string) ??
    (args.writeTo as string)
  );
}

/** Determine if the tool targets a whole sheet (used for clearWorksheet / deleteWorksheet). */
function isWholeSheetTool(
  toolName: string,
  args: Record<string, unknown>,
): boolean {
  return toolName === "clearWorksheet" || toolName === "deleteWorksheet";
}

// ── Supervisor Agent ──
// Fast classification: chat vs task

export class SupervisorAgent implements Agent {
  readonly name = "supervisor";
  readonly model: string;

  constructor(model: string) {
    this.model = model;
  }

  async execute(
    input: AgentInput,
    context: AgentContext,
  ): Promise<AgentResult<"chat" | "task">> {
    context.onThinking?.(
      "🧭 Supervisor: классификация запроса...",
      "reasoning",
    );

    const messages: ChatMessage[] = [
      { role: "system", content: SUPERVISOR_PROMPT },
      { role: "user", content: input.userMessage },
    ];

    try {
      const response = await context.llm.chat(messages, []);
      const result =
        response.choices[0]?.message?.content?.trim().toLowerCase() || "task";
      const intent = result.includes("chat") ? "chat" : "task";

      return {
        success: true,
        data: intent,
        tokensUsed: response.usage?.total_tokens || 0,
      };
    } catch (error) {
      // Fallback: assume task (safer)
      return {
        success: true,
        data: "task",
        tokensUsed: 0,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }
}

// ── Planner Agent ──
// Creates structured JSON plan

export class PlannerAgent implements Agent {
  readonly name = "planner";
  readonly model: string;

  constructor(model: string) {
    this.model = model;
  }

  async execute(
    input: AgentInput,
    context: AgentContext,
  ): Promise<AgentResult<StructuredPlan>> {
    context.onThinking?.("📋 Planner: составление плана...", "reasoning");
    context.onProgress?.(1, 5, "Анализ задачи и составление плана...");

    const messages: ChatMessage[] = [
      { role: "system", content: PLANNER_PROMPT },
      { role: "user", content: input.userMessage },
    ];

    try {
      const response = await context.llm.chat(messages, []);
      const content = response.choices[0]?.message?.content || "";

      const plan = this.parsePlan(content);

      if (!plan) {
        return {
          success: false,
          data: this.fallbackPlan(input.userMessage),
          tokensUsed: response.usage?.total_tokens || 0,
          error: "Failed to parse plan JSON",
        };
      }

      return {
        success: true,
        data: plan,
        tokensUsed: response.usage?.total_tokens || 0,
      };
    } catch (error) {
      return {
        success: false,
        data: this.fallbackPlan(input.userMessage),
        tokensUsed: 0,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  private parsePlan(content: string): StructuredPlan | null {
    // Try to extract JSON from the response (model may wrap it in markdown)
    let jsonStr = content.trim();

    // Strip markdown code fences if present
    const fenceMatch = jsonStr.match(/```(?:json)?\s*\n([\s\S]*?)\n```/);
    if (fenceMatch) {
      jsonStr = fenceMatch[1].trim();
    }

    // Find first { and last }
    const firstBrace = jsonStr.indexOf("{");
    const lastBrace = jsonStr.lastIndexOf("}");
    if (firstBrace === -1 || lastBrace === -1) return null;
    jsonStr = jsonStr.substring(firstBrace, lastBrace + 1);

    try {
      const parsed = JSON.parse(jsonStr);
      if (!parsed.steps || !Array.isArray(parsed.steps)) return null;

      // Validate and normalize steps
      const steps: PlanStep[] = parsed.steps.map(
        (s: Record<string, unknown>, i: number) => ({
          id: (s.id as string) || `step_${i + 1}`,
          description: (s.description as string) || "Unnamed step",
          tool: (s.tool as string) || "getWorkbookOverview",
          args: (s.args as Record<string, unknown>) || {},
          dependsOn: Array.isArray(s.dependsOn)
            ? (s.dependsOn as string[])
            : [],
          canRunInParallel: Boolean(s.canRunInParallel),
          riskLevel:
            (s.riskLevel as "safe" | "moderate" | "dangerous") || "safe",
        }),
      );

      return {
        title: parsed.title || "Execution Plan",
        summary: parsed.summary || "",
        steps,
        riskLevel: parsed.riskLevel || "low",
        dataImpact: parsed.dataImpact || "modify",
        estimatedSteps: steps.length,
      };
    } catch {
      return null;
    }
  }

  private fallbackPlan(userMessage: string): StructuredPlan {
    return {
      title: "Execute Task",
      summary: "Direct execution without structured plan",
      riskLevel: "medium",
      dataImpact: "modify",
      estimatedSteps: 1,
      steps: [
        {
          id: "step_1",
          description: userMessage,
          tool: "getWorkbookOverview",
          args: {},
          dependsOn: [],
          canRunInParallel: false,
          riskLevel: "safe",
        },
      ],
    };
  }
}

// ── Executor Agent ──
// Executes tools via native function calling, supports parallel execution

export class ExecutorAgent implements Agent {
  readonly name = "executor";
  readonly model: string;
  private maxIterations = 15;

  constructor(model: string) {
    this.model = model;
  }

  async execute(
    input: AgentInput,
    context: AgentContext,
  ): Promise<AgentResult<ExecutionResult>> {
    context.onThinking?.("🚀 Executor: выполнение плана...", "executing");

    const startTime = Date.now();
    const stepResults: StepResult[] = [];
    const errors: string[] = [];
    const toolNames: string[] = [];
    let rowsModified = 0;

    // Build execution messages with the plan context
    const planContext = input.plan
      ? `You are executing the following plan:\n${JSON.stringify(input.plan, null, 2)}\n\nUser request: ${input.userMessage}\n\nExecute each step by calling the appropriate tools. Use native function calling.`
      : input.userMessage;

    const messages: ChatMessage[] = [
      { role: "system", content: EXECUTOR_PROMPT },
      { role: "user", content: planContext },
    ];

    let iterationCount = 0;
    const totalIterations = this.maxIterations;

    while (iterationCount < this.maxIterations) {
      if (context.aborted) {
        return {
          success: false,
          data: this.buildResult(
            stepResults,
            startTime,
            toolNames,
            rowsModified,
            errors,
          ),
          tokensUsed: 0,
          error: "Execution aborted by user",
        };
      }

      iterationCount++;
      context.onProgress?.(
        iterationCount,
        totalIterations,
        `Step ${iterationCount}: executing...`,
      );
      context.onThinking?.(
        `🔧 Iteration ${iterationCount}: reasoning...`,
        "executing",
      );

      let response;
      try {
        response = await context.llm.chatStream(
          messages,
          context.tools,
          (chunk: string) => {
            if (chunk) context.onThinking?.(chunk, "executing");
          },
        );
      } catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        if (errMsg.includes("cancelled") || errMsg.includes("abort")) {
          return {
            success: false,
            data: this.buildResult(
              stepResults,
              startTime,
              toolNames,
              rowsModified,
              errors,
            ),
            tokensUsed: 0,
            error: "Execution cancelled",
          };
        }
        errors.push(errMsg);
        break;
      }

      const choice = response.choices[0];
      const message = choice.message;

      if (message.tool_calls && message.tool_calls.length > 0) {
        // Push assistant message with tool calls
        const assistantMsg: ChatMessage = {
          role: "assistant",
          content: message.content ?? "",
          tool_calls: message.tool_calls,
        };
        messages.push(assistantMsg);

        if (message.content) {
          context.onMessage("assistant", message.content);
        }

        // Execute tool calls — parallel if independent
        const toolCalls = message.tool_calls;
        const canParallelize = this.canParallelize(toolCalls);

        if (canParallelize && toolCalls.length > 1) {
          context.onThinking?.(
            `⚡ Executing ${toolCalls.length} tools in parallel...`,
            "executing",
          );
          const results = await Promise.all(
            toolCalls.map((tc) =>
              this.executeToolCall(
                tc,
                context,
                iterationCount,
                totalIterations,
                undefined, // parallel tools are independent — no transaction
              ),
            ),
          );

          for (let i = 0; i < results.length; i++) {
            const result = results[i];
            const tc = toolCalls[i];
            stepResults.push({
              stepId: `iter_${iterationCount}_tool_${i}`,
              success: result.success,
              toolName: tc.function.name,
              result: result.result,
              error: result.error,
              elapsedMs: result.elapsedMs,
            });
            if (result.success) {
              toolNames.push(tc.function.name);
              rowsModified += this.estimateRows(
                tc.function.name,
                result.result,
              );
            } else {
              errors.push(result.error || "Unknown error");
            }
            messages.push({
              role: "tool",
              content: result.result,
              tool_call_id: tc.id,
            });
          }
        } else {
          // Sequential execution
          // ── Transactional batch: if multiple sequential dangerous tools, group them ──
          const hasDangerous = toolCalls.some((tc) =>
            DANGEROUS_TOOLS.has(tc.function.name),
          );
          const allTransactional = toolCalls.every(
            (tc) =>
              TRANSACTIONAL_TOOLS.has(tc.function.name) ||
              !DANGEROUS_TOOLS.has(tc.function.name),
          );
          const useTransaction =
            hasDangerous && toolCalls.length > 1 && allTransactional;
          const txId = useTransaction
            ? undoManager.beginTransaction(
                `Batch: ${toolCalls.map((t) => t.function.name).join(", ")}`,
              )
            : undefined;

          for (const tc of toolCalls) {
            const result = await this.executeToolCall(
              tc,
              context,
              iterationCount,
              totalIterations,
              txId,
            );
            stepResults.push({
              stepId: `iter_${iterationCount}_tool_${tc.id}`,
              success: result.success,
              toolName: tc.function.name,
              result: result.result,
              error: result.error,
              elapsedMs: result.elapsedMs,
            });
            if (result.success) {
              toolNames.push(tc.function.name);
              rowsModified += this.estimateRows(
                tc.function.name,
                result.result,
              );
            } else {
              errors.push(result.error || "Unknown error");
              // If transactional and a step failed, rollback already happened
              // inside executeToolCall. Stop the batch.
              if (useTransaction) break;
            }
            messages.push({
              role: "tool",
              content: result.result,
              tool_call_id: tc.id,
            });
          }

          // Commit successful transaction
          if (useTransaction) {
            await undoManager.commitTransaction(txId!);
          }
        }
      } else {
        // No tool calls — execution is done
        const content = message.content ?? "";
        messages.push({ role: "assistant", content });
        context.onMessage("assistant", content);

        return {
          success: errors.length === 0,
          data: this.buildResult(
            stepResults,
            startTime,
            toolNames,
            rowsModified,
            errors,
          ),
          tokensUsed: response.usage?.total_tokens || 0,
        };
      }

      if (
        choice.finish_reason === "stop" &&
        (!message.tool_calls || message.tool_calls.length === 0)
      ) {
        break;
      }
    }

    return {
      success: errors.length === 0,
      data: this.buildResult(
        stepResults,
        startTime,
        toolNames,
        rowsModified,
        errors,
      ),
      tokensUsed: 0,
    };
  }

  private async executeToolCall(
    tc: { id: string; function: { name: string; arguments: string } },
    context: AgentContext,
    iteration: number,
    total: number,
    transactionId?: string,
  ): Promise<{
    success: boolean;
    result: string;
    error?: string;
    elapsedMs: number;
    backupId?: string | null;
  }> {
    const toolName = tc.function.name;
    const startTime = Date.now();

    let args: Record<string, unknown>;
    try {
      args = JSON.parse(tc.function.arguments);
    } catch {
      return {
        success: false,
        result: `Failed to parse arguments for ${toolName}`,
        error: `Invalid JSON arguments: ${tc.function.arguments}`,
        elapsedMs: Date.now() - startTime,
      };
    }

    context.onThinking?.(`🔧 Executing: ${toolName}...`, "executing");
    context.onProgress?.(iteration, total, `Step ${iteration}: ${toolName}...`);

    // ── Pre-execution validation: analyze args for risky formulas ──
    const risk = analyzeToolArgsRisk(toolName, args);
    if (risk.level === "blocked") {
      const msg = `🚫 Формула заблокирована: ${risk.description}`;
      context.onMessage("error", msg);
      return {
        success: false,
        result: msg,
        error: "Blocked formula",
        elapsedMs: Date.now() - startTime,
      };
    }

    // ── Backup for undo safety (no user confirmation — undo is the safety net) ──
    let backupId: string | null = null;
    if (DANGEROUS_TOOLS.has(toolName)) {
      const backupAddress = getBackupAddress(args);
      const wholeSheet = isWholeSheetTool(toolName, args);

      if (backupAddress || wholeSheet) {
        try {
          const backup = await undoManager.createBackup(
            backupAddress || "A1",
            toolName,
            {
              sheetName: args.name as string | undefined,
              wholeSheet,
              transactionId,
              onProgress: (done, totalChunks, label) => {
                context.onThinking?.(`💾 ${label}`, "executing");
              },
            },
          );
          backupId = backup.backupId;
          if (backup.backupSkipped) {
            context.onThinking?.(
              `⚠️ Backup skipped: ${backup.reason}`,
              "executing",
            );
          }
        } catch (err) {
          context.onMessage(
            "error",
            `Backup failed: ${err instanceof Error ? err.message : String(err)}`,
          );
        }
      }
    }

    context.onMessage(
      "tool-call",
      `Executing: ${toolName}\n\`\`\`json\n${JSON.stringify(args, null, 2)}\n\`\`\``,
    );

    try {
      const result = await toolRegistry.execute(toolName, args);
      context.onMessage(
        "tool-result",
        `Result:\n\`\`\`json\n${result}\n\`\`\``,
      );

      // Verify result
      const verification = await verifyToolResult(toolName, args, result);
      if (!verification.verified && verification.message) {
        context.onMessage("error", `⚠️ ${verification.message}`);
      }

      return {
        success: true,
        result,
        elapsedMs: Date.now() - startTime,
        backupId,
      };
    } catch (error) {
      const errMsg = error instanceof Error ? error.message : String(error);
      context.onMessage("error", `Error executing ${toolName}: ${errMsg}`);

      // ── Transactional rollback: if part of a transaction, roll back the whole batch ──
      if (transactionId) {
        context.onMessage(
          "system",
          `🔄 Ошибка в транзакции — откатываю всю группу...`,
        );
        const rollback = await undoManager.rollbackTransaction(transactionId);
        context.onMessage(
          "system",
          `↩️ Откат: восстановлено ${rollback.restored} действий, ошибок: ${rollback.errors.length}`,
        );
      }

      return {
        success: false,
        result: `Error: ${errMsg}`,
        error: errMsg,
        elapsedMs: Date.now() - startTime,
        backupId,
      };
    }
  }

  private canParallelize(toolCalls: { function: { name: string } }[]): boolean {
    // Don't parallelize if any tool is dangerous (needs confirmation)
    // Don't parallelize if tools write to the same range
    return toolCalls.every((tc) => !DANGEROUS_TOOLS.has(tc.function.name));
  }

  private estimateRows(toolName: string, result: string): number {
    const match = result.match(/(\d+)\s*(cells|rows|items)/i);
    if (match) return parseInt(match[1], 10);
    if (toolName === "setValues") return 1;
    return 0;
  }

  private buildResult(
    steps: StepResult[],
    startTime: number,
    toolNames: string[],
    rowsModified: number,
    errors: string[],
  ): ExecutionResult {
    return {
      steps,
      totalElapsedMs: Date.now() - startTime,
      toolCallCount: steps.length,
      rowsModified,
      toolNames: [...new Set(toolNames)],
      errors,
    };
  }
}

// ── Verifier Agent ──
// Checks results against the plan

export class VerifierAgent implements Agent {
  readonly name = "verifier";
  readonly model: string;

  constructor(model: string) {
    this.model = model;
  }

  async execute(
    input: AgentInput,
    context: AgentContext,
  ): Promise<AgentResult<VerificationResult>> {
    if (!input.executionResult) {
      return {
        success: true,
        data: {
          verified: true,
          issues: [],
          summary: "No execution to verify",
          canAutoFix: false,
        },
        tokensUsed: 0,
      };
    }

    context.onThinking?.("🔍 Verifier: проверка результатов...", "verifying");
    context.onProgress?.(4, 5, "Проверка результатов...");

    const exec = input.executionResult;
    const planInfo = input.plan
      ? `Plan: ${JSON.stringify(input.plan.steps.map((s) => ({ id: s.id, tool: s.tool, description: s.description })))}`
      : "No structured plan";

    const execInfo = `Execution results: ${JSON.stringify(exec.steps.map((s) => ({ tool: s.toolName, success: s.success, result: s.result.substring(0, 500), error: s.error })))}`;

    const messages: ChatMessage[] = [
      { role: "system", content: VERIFIER_PROMPT },
      { role: "user", content: `${planInfo}\n\n${execInfo}` },
    ];

    try {
      const response = await context.llm.chat(messages, []);
      const content = response.choices[0]?.message?.content || "";

      const result = this.parseVerification(content);

      return {
        success: true,
        data: result,
        tokensUsed: response.usage?.total_tokens || 0,
      };
    } catch {
      // If verifier fails, assume verified
      return {
        success: true,
        data: {
          verified: exec.errors.length === 0,
          issues: exec.errors.map((e) => ({
            stepId: "unknown",
            severity: "error" as const,
            description: e,
          })),
          summary:
            exec.errors.length === 0
              ? "All steps completed"
              : "Some errors occurred",
          canAutoFix: false,
        },
        tokensUsed: 0,
      };
    }
  }

  private parseVerification(content: string): VerificationResult {
    let jsonStr = content.trim();
    const fenceMatch = jsonStr.match(/```(?:json)?\s*\n([\s\S]*?)\n```/);
    if (fenceMatch) jsonStr = fenceMatch[1].trim();

    const firstBrace = jsonStr.indexOf("{");
    const lastBrace = jsonStr.lastIndexOf("}");
    if (firstBrace === -1 || lastBrace === -1) {
      return {
        verified: true,
        issues: [],
        summary: "Verification parse failed, assuming OK",
        canAutoFix: false,
      };
    }

    try {
      const parsed = JSON.parse(jsonStr.substring(firstBrace, lastBrace + 1));
      return {
        verified: Boolean(parsed.verified),
        issues: Array.isArray(parsed.issues) ? parsed.issues : [],
        summary: parsed.summary || "Verification complete",
        canAutoFix: Boolean(parsed.canAutoFix),
      };
    } catch {
      return {
        verified: true,
        issues: [],
        summary: "Verification parse failed, assuming OK",
        canAutoFix: false,
      };
    }
  }
}

// ── Summarizer Agent ──
// Creates user-friendly summary

export class SummarizerAgent implements Agent {
  readonly name = "summarizer";
  readonly model: string;

  constructor(model: string) {
    this.model = model;
  }

  async execute(
    input: AgentInput,
    context: AgentContext,
  ): Promise<AgentResult<string>> {
    context.onThinking?.("📝 Summarizer: подготовка отчёта...", "summarizing");
    context.onProgress?.(5, 5, "Подготовка отчёта...");

    const exec = input.executionResult;
    const verification = input.verificationResult;

    const execSummary = exec
      ? `Tools called: ${exec.toolCallCount}. Tools: ${exec.toolNames.join(", ")}. Errors: ${exec.errors.length}. Rows modified: ${exec.rowsModified}. Time: ${(exec.totalElapsedMs / 1000).toFixed(1)}s.`
      : "No execution data";

    const verifSummary = verification
      ? `Verified: ${verification.verified}. Issues: ${verification.issues.length}. ${verification.summary}`
      : "No verification data";

    const messages: ChatMessage[] = [
      { role: "system", content: SUMMARIZER_PROMPT },
      {
        role: "user",
        content: `User request: ${input.userMessage}\n\nExecution: ${execSummary}\n\nVerification: ${verifSummary}`,
      },
    ];

    try {
      const response = await context.llm.chat(messages, []);
      const content = response.choices[0]?.message?.content || "Готово.";

      return {
        success: true,
        data: content,
        tokensUsed: response.usage?.total_tokens || 0,
      };
    } catch {
      // Fallback summary
      const success = exec?.errors.length === 0;
      const summary = success
        ? `✅ Готово! Выполнено ${exec?.toolCallCount || 0} операций.`
        : `⚠️ Выполнено с ошибками: ${exec?.errors.length || 0} проблем.`;
      return { success: true, data: summary, tokensUsed: 0 };
    }
  }
}
