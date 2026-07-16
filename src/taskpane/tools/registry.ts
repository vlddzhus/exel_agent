/**
 * registry.ts — ЕДИНЫЙ реестр инструментов Excel-агента.
 *
 * См. docs/03-TOOLS-SPEC.md §0.3, §7.
 *
 * Предлагает два способа регистрации:
 *
 * 1. НОВЫЙ (рекомендован для всех новых инструментов Фазы 1):
 *    defineTool({...}) — декларативный, с riskLevel/requiresUndo/estimateCells.
 *    пример в docs/03-TOOLS-SPEC.md §0.3.
 *
 * 2. LEGACY (для существующих инструментов до их миграции):
 *    toolRegistry.register(name, description, parameters, fn, requiresConfirmation)
 *
 * Источник истины riskLevel — defineTool. Дублирующий DANGEROUS_TOOLS в agents.ts
 * будет удалён после миграции всех инструментов на defineTool.
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type RiskLevel = "safe" | "moderate" | "dangerous";

/**
 * Контекст выполнения инструмента (будет расширяться: undo coordinator, logger).
 * В Фазе 1 пока используется минимальная версия — сами инструменты вызывают
 * undoManager напрямую. В Фазе 2+ это станет полноценным ToolContext.
 */
export interface ToolContext {
  /** Идентификатор запроса для корреляции с undo/логами. */
  requestId?: string;
}

/**
 * Стандартизованный результат инструмента.
 * Каждый инструмент ОБЯЗАН возвращать ToolResult (см. docs/03-TOOLS-SPEC.md §6.2).
 */
export interface ToolResult {
  ok: boolean;
  /** Человекочитаемое summary для LLM и UI. */
  summary: string;
  /** Структурированный ответ для LLM (опционально). */
  data?: unknown;
  /** Сколько ячеек затронуто (для статистики). */
  cellsAffected?: number;
  /** Детали ошибки, если ok=false. */
  error?: {
    code: string;
    message: string;
    retryable: boolean;
  };
}

/**
 * Полное определение инструмента (новый API).
 */
export interface ToolDefinition {
  name: string;
  /** Описание на русском, объясняет когда использовать (для LLM system prompt). */
  description: string;
  /** Zod-схема или JSON Schema параметров. */
  parameters: Record<string, unknown>;
  riskLevel: RiskLevel;
  /** Требуется ли undo-снапшот перед выполнением. */
  requiresUndo: boolean;
  /** Оценка количества затрагиваемых ячеек (для лимитов и прогресса). */
  estimateCells: (args: Record<string, unknown>) => number;
  /** Функция выполнения. */
  execute: (
    args: Record<string, unknown>,
    ctx?: ToolContext,
  ) => Promise<ToolResult>;
}

// Legacy types (для обратной совместимости)
export interface ToolParam {
  type: string;
  description?: string;
  enum?: string[];
  items?: Record<string, unknown>;
  properties?: Record<string, ToolParam>;
  required?: string[];
}

export interface ToolSchema {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
}

export type ToolFunction = (args: Record<string, unknown>) => Promise<string>;

interface ToolEntry {
  schema: ToolSchema;
  fn: ToolFunction;
  // Новые поля (undefined для legacy-инструментов)
  definition?: ToolDefinition;
  riskLevel?: RiskLevel;
  requiresUndo?: boolean;
  /** Legacy: заменяется на riskLevel === "dangerous" в новых инструментах. */
  requiresConfirmation: boolean;
}

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

class ToolRegistry {
  private tools = new Map<string, ToolEntry>();

  /**
   * НОВЫЙ API: регистрация через defineTool-объект.
   * Используйте для всех новых инструментов Фазы 1.
   */
  registerDefinition(def: ToolDefinition): void {
    if (this.tools.has(def.name)) {
      throw new Error(`Tool already registered: ${def.name}`);
    }
    const fn: ToolFunction = async (args) => {
      const result = await def.execute(args);
      return JSON.stringify(result);
    };
    this.tools.set(def.name, {
      schema: {
        name: def.name,
        description: def.description,
        parameters: def.parameters,
      },
      fn,
      definition: def,
      riskLevel: def.riskLevel,
      requiresUndo: def.requiresUndo,
      // Новые "dangerous" инструменты требуют подтверждения (legacy-флаг)
      requiresConfirmation: def.riskLevel === "dangerous",
    });
  }

  /**
   * LEGACY API: оставить для существующих инструментов до их миграции.
   * Будет удалён после полного перехода на defineTool.
   */
  register(
    name: string,
    description: string,
    parameters: Record<string, unknown>,
    fn: ToolFunction,
    requiresConfirmation = false,
  ): void {
    if (this.tools.has(name)) {
      throw new Error(`Tool already registered: ${name}`);
    }
    this.tools.set(name, {
      schema: { name, description, parameters },
      fn,
      requiresConfirmation,
      // Legacy-инструменты без явного riskLevel — по умолчанию moderate
      riskLevel: requiresConfirmation ? "dangerous" : "moderate",
    });
  }

  getTool(name: string): ToolEntry | undefined {
    return this.tools.get(name);
  }

  /**
   * Получить полное определение инструмента (новый API).
   * undefined для legacy-инструментов без definition.
   */
  getDefinition(name: string): ToolDefinition | undefined {
    return this.tools.get(name)?.definition;
  }

  getSchemas() {
    return Array.from(this.tools.values()).map((t) => ({
      type: "function" as const,
      function: t.schema,
    }));
  }

  getToolNames(): string[] {
    return Array.from(this.tools.keys());
  }

  /** Все зарегистрированные определения (только новый API). */
  getDefinitions(): ToolDefinition[] {
    return Array.from(this.tools.values())
      .map((t) => t.definition)
      .filter((d): d is ToolDefinition => d !== undefined);
  }

  /** Legacy: требует ли подтверждения. */
  requiresConfirmation(name: string): boolean {
    return this.tools.get(name)?.requiresConfirmation ?? false;
  }

  /** НОВОЕ: risk level инструмента (единый источник истины). */
  riskLevel(name: string): RiskLevel {
    return this.tools.get(name)?.riskLevel ?? "moderate";
  }

  /** НОВОЕ: требует ли undo-снапшота перед выполнением. */
  requiresUndo(name: string): boolean {
    return this.tools.get(name)?.requiresUndo ?? false;
  }

  /** НОВОЕ: оценка количества ячеек для инструмента. */
  estimateCells(name: string, args: Record<string, unknown>): number {
    const def = this.tools.get(name)?.definition;
    if (!def) return 0;
    try {
      return def.estimateCells(args);
    } catch {
      return 0;
    }
  }

  async execute(
    name: string,
    args: Record<string, unknown>,
    ctx?: ToolContext,
  ): Promise<string> {
    const tool = this.tools.get(name);
    if (!tool) {
      throw new Error(`Unknown tool: ${name}`);
    }
    // Если есть definition — вызываем через него (для контекста/логирования)
    if (tool.definition) {
      const result = await tool.definition.execute(args, ctx);
      return JSON.stringify(result);
    }
    return tool.fn(args);
  }
}

export const toolRegistry = new ToolRegistry();

// ---------------------------------------------------------------------------
// Удобный helper: defineTool — просто возвращает определение как есть.
// Использовать так:
//   export const myTool = defineTool({...});
//   toolRegistry.registerDefinition(myTool);
// ---------------------------------------------------------------------------

export function defineTool<T extends ToolDefinition>(def: T): T {
  return def;
}
