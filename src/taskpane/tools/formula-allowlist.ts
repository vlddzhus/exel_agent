/**
 * Formula Allowlist + Pre-Execution Validation Layer.
 *
 * Classifies Excel functions used in a formula into safe / blocked / confirm-required,
 * and produces a human-readable risk description for the user.
 *
 * Blocked functions (never executed):
 *   WEBSERVICE, FILTERXML, CALL, REGISTER.ID, EXEC,
 *   and Excel 4.0 macro functions (REGISTER, UNREGISTER, ARGUMENT, RESULT, etc.)
 *
 * Confirm-required:
 *   HYPERLINK (can leak data via external URLs / phishing)
 *
 * Everything else is allowed (Excel has thousands of functions; allowlist is open
 * with a blocked-list, plus a small explicit safe set for fast-path validation).
 */

// ---------------------------------------------------------------------------
// Function classification
// ---------------------------------------------------------------------------

/** Functions that MUST NEVER be executed — they exfiltrate data or run code. */
export const BLOCKED_FUNCTIONS: ReadonlySet<string> = new Set([
  // Network / data exfiltration
  "WEBSERVICE",
  "FILTERXML",
  // Dynamic code / legacy macro invocation
  "CALL",
  "REGISTER.ID",
  "EXEC",
  // Excel 4.0 (XLM) macro functions
  "REGISTER",
  "UNREGISTER",
  "ARGUMENT",
  "RESULT",
  "FORMULA",
  "EXECUTE",
  "RUN",
  "GOTO",
  "DEFINE.NAME",
  "SET.NAME",
  "WORKBOOK",
  "DOCUMENTS",
  "FILES",
  "DIRECTORY",
  "ALERT",
  "MESSAGE",
  "EVALUATE",
  "CHAR",
]);

/**
 * HYPERLINK is special: it's commonly useful but can be used for data exfiltration
 * and phishing. It requires explicit user confirmation.
 */
export const CONFIRM_REQUIRED_FUNCTIONS: ReadonlySet<string> = new Set([
  "HYPERLINK",
]);

/**
 * A small explicit safe set for fast-path validation. Not exhaustive —
 * any function NOT in BLOCKED or CONFIRM_REQUIRED is allowed by default.
 */
export const SAFE_FUNCTIONS: ReadonlySet<string> = new Set([
  // Math
  "SUM",
  "AVERAGE",
  "COUNT",
  "COUNTA",
  "MAX",
  "MIN",
  "ROUND",
  "ROUNDUP",
  "ROUNDDOWN",
  "SUMIF",
  "SUMIFS",
  "SUMPRODUCT",
  "PRODUCT",
  "POWER",
  "SQRT",
  "MOD",
  "INT",
  "ABS",
  "CEILING",
  "FLOOR",
  "TRUNC",
  "RAND",
  "RANDBETWEEN",
  "LOG",
  "LN",
  "EXP",
  "PI",
  // Lookup
  "VLOOKUP",
  "HLOOKUP",
  "INDEX",
  "MATCH",
  "LOOKUP",
  "XLOOKUP",
  "XMATCH",
  "OFFSET",
  "INDIRECT",
  "CHOOSE",
  "ROW",
  "COLUMN",
  "ROWS",
  "COLUMNS",
  "ADDRESS",
  "AREAS",
  // Logical
  "IF",
  "AND",
  "OR",
  "NOT",
  "IFERROR",
  "IFNA",
  "IFS",
  "SWITCH",
  "TRUE",
  "FALSE",
  // Text
  "CONCAT",
  "CONCATENATE",
  "TEXTJOIN",
  "LEFT",
  "RIGHT",
  "MID",
  "LEN",
  "FIND",
  "SEARCH",
  "REPLACE",
  "SUBSTITUTE",
  "TEXT",
  "TRIM",
  "UPPER",
  "LOWER",
  "PROPER",
  "VALUE",
  "REPT",
  // Date/Time
  "TODAY",
  "NOW",
  "DATE",
  "DATEDIF",
  "DAY",
  "MONTH",
  "YEAR",
  "WEEKDAY",
  "EOMONTH",
  "EDATE",
  "WORKDAY",
  "NETWORKDAYS",
  "HOUR",
  "MINUTE",
  "SECOND",
  "TIME",
  // Statistical
  "STDEV",
  "STDEVP",
  "STDEV.S",
  "STDEV.P",
  "VAR",
  "VARP",
  "MEDIAN",
  "MODE",
  "QUARTILE",
  "PERCENTILE",
  "RANK",
  "LARGE",
  "SMALL",
  "CORREL",
  "COVAR",
  "NORM.DIST",
  "NORM.INV",
  // Russian equivalents
  "СУММ",
  "СРЗНАЧ",
  "СЧЁТ",
  "СЧЁТЗ",
  "МАКС",
  "МИН",
  "ОКРУГЛ",
  "СУММПРОИЗВ",
  "СУММЕСЛИ",
  "СУММЕСЛИМН",
  "ВПР",
  "ГПР",
  "ИНДЕКС",
  "ПОИСКПОЗ",
  "ПРОСМОТРХ",
  "ЕСЛИ",
  "И",
  "ИЛИ",
  "НЕ",
  "ЕСЛИОШИБКА",
  "ЕСНД",
  "ЕСЛИМН",
  "ПЕРЕКЛ",
  "СЦЕПИТЬ",
  "ЛЕВСИМВ",
  "ПРАВСИМВ",
  "ПСТР",
  "ДЛСТР",
  "НАЙТИ",
  "ЗАМЕНИТЬ",
  "ПОДСТАВИТЬ",
  "ТЕКСТ",
  "ОБЪЕДИНИТЬ",
  "СЕГОДНЯ",
  "ТДАТА",
  "ДАТА",
  "РАЗНДАТ",
  "ДЕНЬ",
  "МЕСЯЦ",
  "ГОД",
  "СТАНДОТКЛОН",
  "МЕДИАНА",
  "КВАРТИЛЬ",
  "СТАНДОТКЛОНП",
  "ДИСП",
  "СТЕПЕНЬ",
  "КОРЕНЬ",
  "ОСТАТ",
  "ЦЕЛОЕ",
  "ПРОИЗВЕД",
]);

// ---------------------------------------------------------------------------
// Risk report
// ---------------------------------------------------------------------------

export type RiskLevel = "safe" | "confirm" | "blocked";

export interface FunctionRisk {
  name: string;
  level: RiskLevel;
  reason: string;
}

export interface FormulaRiskReport {
  /** Overall risk level = worst of all detected functions. */
  level: RiskLevel;
  /** Per-function risk breakdown. */
  functions: FunctionRisk[];
  /** Human-readable summary in Russian. */
  description: string;
  /** True if formula contains any function call at all. */
  hasFunctions: boolean;
  /** Raw list of all function names detected (uppercased). */
  detected: string[];
}

// ---------------------------------------------------------------------------
// Function name extraction
// ---------------------------------------------------------------------------

/**
 * Extract all function-call names from a formula string.
 * Matches identifiers (including "." for names like REGISTER.ID) immediately
 * followed by "(". Uses negative lookbehind so nested calls are all captured.
 *
 * Examples:
 *   "SUM(A1:B2)"               → ["SUM"]
 *   "IF(A1>0, VLOOKUP(B1,C:D,2,0), 0)" → ["IF", "VLOOKUP"]
 *   "A1+B2"                    → []
 *   "СУММ(A1:A10)"             → ["СУММ"]
 *   "ROUND(SUM(A1), 2) & TEXT(NOW(), \"Y\")" → ["ROUND", "SUM", "TEXT", "NOW"]
 *   "REGISTER.ID(\"x\", \"y\")" → ["REGISTER.ID"]
 */
export function extractFunctionNames(formula: string): string[] {
  if (!formula) return [];
  // Negative lookbehind: function name must NOT be preceded by an identifier char.
  // This avoids matching inside cell refs (A1SUM doesn't happen, but be safe)
  // and allows nested calls (after "(", the next identifier is a new function).
  // "." is allowed in the name body for REGISTER.ID-style names.
  const re =
    /(?<![A-Za-zА-Яа-я0-9_])([A-Za-zА-Яа-я_][A-Za-zА-Яа-я0-9_.]*)\s*\(/g;
  const names: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(formula)) !== null) {
    names.push(m[1].toUpperCase());
  }
  return names;
}

// ---------------------------------------------------------------------------
// Risk analysis
// ---------------------------------------------------------------------------

const REASONS: Record<string, string> = {
  WEBSERVICE:
    "Выполняет HTTP-запрос на внешний URL — может утечь данные листа.",
  FILTERXML:
    "Парсит XML из произвольного URL — используется с WEBSERVICE для эксфильтрации.",
  CALL: "Вызывает процедуры из DLL — удалённое выполнение кода.",
  "REGISTER.ID":
    "Регистрирует и вызывает внешние DLL-функции — удалённое выполнение кода.",
  EXEC: "Запускает внешние программы — удалённое выполнение кода.",
  REGISTER: "Excel 4.0 macro: регистрирует DLL-функции — выполнение кода.",
  UNREGISTER: "Excel 4.0 macro: управление DLL-регистрацией.",
  ARGUMENT: "Excel 4.0 macro function.",
  RESULT: "Excel 4.0 macro function.",
  FORMULA: "Excel 4.0 macro: динамически создаёт формулы.",
  EXECUTE: "Excel 4.0 macro: запускает команды.",
  RUN: "Excel 4.0 macro: вызывает макросы.",
  GOTO: "Excel 4.0 macro: управление потоком.",
  "DEFINE.NAME": "Excel 4.0 macro: создаёт именованные диапазоны.",
  "SET.NAME": "Excel 4.0 macro: изменяет именованные диапазоны.",
  WORKBOOK: "Excel 4.0 macro: доступ к книгам.",
  DOCUMENTS: "Excel 4.0 macro: перечисляет документы.",
  FILES: "Excel 4.0 macro: перечисляет файлы на диске.",
  DIRECTORY: "Excel 4.0 macro: доступ к файловой системе.",
  ALERT: "Excel 4.0 macro: показывает диалоги.",
  MESSAGE: "Excel 4.0 macro: показывает сообщения.",
  EVALUATE: "Excel 4.0 macro: вычисляет произвольное выражение.",
  CHAR: "Может использоваться для обфускации вредоносных формул (CHAR-инъекция).",
  HYPERLINK:
    "Создаёт кликабельную ссылку на внешний URL — может вести на фишинг или утекать данные.",
};

/**
 * Analyze a formula for risky functions.
 *
 * @param formula — formula WITHOUT leading "="
 */
export function analyzeFormulaRisk(formula: string): FormulaRiskReport {
  const detected = extractFunctionNames(formula);
  const hasFunctions = detected.length > 0;

  if (!hasFunctions) {
    return {
      level: "safe",
      functions: [],
      description: "Формула не содержит вызовов функций.",
      hasFunctions: false,
      detected: [],
    };
  }

  const functions: FunctionRisk[] = [];
  const seen = new Set<string>();
  let worst: RiskLevel = "safe";

  for (const name of detected) {
    if (seen.has(name)) continue;
    seen.add(name);

    let level: RiskLevel = "safe";
    let reason = "Безопасная функция.";

    if (BLOCKED_FUNCTIONS.has(name)) {
      level = "blocked";
      reason = REASONS[name] || "Заблокированная функция.";
      if (worst !== "blocked") worst = "blocked";
    } else if (CONFIRM_REQUIRED_FUNCTIONS.has(name)) {
      level = "confirm";
      reason = REASONS[name] || "Требует подтверждения.";
      if (worst === "safe") worst = "confirm";
    }

    functions.push({ name, level, reason });
  }

  const blockedNames = functions
    .filter((f) => f.level === "blocked")
    .map((f) => f.name);
  const confirmNames = functions
    .filter((f) => f.level === "confirm")
    .map((f) => f.name);

  let description: string;
  if (worst === "blocked") {
    description =
      `🚫 Заблокированные функции: ${blockedNames.join(", ")}. ` +
      functions
        .filter((f) => f.level === "blocked")
        .map((f) => f.reason)
        .join(" ");
  } else if (worst === "confirm") {
    description =
      `⚠️ Требует подтверждения: ${confirmNames.join(", ")}. ` +
      functions
        .filter((f) => f.level === "confirm")
        .map((f) => f.reason)
        .join(" ");
  } else {
    description = `✅ Все функции безопасны: ${detected.join(", ")}.`;
  }

  return { level: worst, functions, description, hasFunctions: true, detected };
}

// ---------------------------------------------------------------------------
// Pre-execution validation for tool args
// ---------------------------------------------------------------------------

export interface ToolRiskReport {
  level: RiskLevel;
  description: string;
  /** Formula risk reports for any formula args found. */
  formulaReports: FormulaRiskReport[];
}

/**
 * Pre-execution validation layer: inspect tool args for risky formulas.
 *
 * Scans common formula-bearing args: `formula`, `values` (2D array — any string
 * cell starting with "="), and returns a combined risk report.
 *
 * @param toolName — tool being invoked
 * @param args     — parsed tool arguments
 */
export function analyzeToolArgsRisk(
  toolName: string,
  args: Record<string, unknown>,
): ToolRiskReport {
  const formulaReports: FormulaRiskReport[] = [];

  // 1. Direct `formula` arg (setFormula, fillFormula)
  if (typeof args.formula === "string") {
    const f = (args.formula as string).replace(/^=/, "").trim();
    if (f) formulaReports.push(analyzeFormulaRisk(f));
  }

  // 2. `values` 2D array — any string cell starting with "=" is a formula
  if (Array.isArray(args.values)) {
    for (const row of args.values as unknown[]) {
      if (!Array.isArray(row)) continue;
      for (const cell of row as unknown[]) {
        if (typeof cell === "string" && cell.startsWith("=")) {
          formulaReports.push(analyzeFormulaRisk(cell.slice(1)));
        }
      }
    }
  }

  // 3. HYPERLINK in setValues is already escaped by range-tools (prefixed with '),
  //    but if a user explicitly passes a formula via setFormula, we catch it here.

  let worst: RiskLevel = "safe";
  for (const r of formulaReports) {
    if (r.level === "blocked") {
      worst = "blocked";
      break;
    }
    if (r.level === "confirm" && worst === "safe") worst = "confirm";
  }

  let description: string;
  if (worst === "blocked") {
    description = formulaReports
      .filter((r) => r.level === "blocked")
      .map((r) => r.description)
      .join(" ");
  } else if (worst === "confirm") {
    description = formulaReports
      .filter((r) => r.level === "confirm")
      .map((r) => r.description)
      .join(" ");
  } else if (formulaReports.length > 0) {
    description = "Формулы проверены — рисков не обнаружено.";
  } else {
    description = `Инструмент ${toolName}: формулы отсутствуют — риск минимальный.`;
  }

  return { level: worst, description, formulaReports };
}
