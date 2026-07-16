/**
 * read.ts — Read-инструменты агента (категория R в docs/03-TOOLS-SPEC.md §1).
 *
 * Принципы (см. docs/03-TOOLS-SPEC.md §0):
 *   - Каждый инструмент = defineTool с riskLevel/requiresUndo/estimateCells.
 *   - Возвращает ToolResult с summary на русском (для LLM/UI).
 *   - Валидация через _shared/validation, адреса через _shared/address.
 *   - Батчинг: один context.sync() на пакет операций (никогда в цикле).
 *
 * Инструменты:
 *   R1 getWorkbookOverview — обзор книги, ОДИН sync (фикс N+1).
 *   R2 getRange            — содержимое диапазона с типами.
 *   R3 getRangeStats       — min/max/sum/avg/count/unique в JS.
 *   R4 detectDataTypes     — типы колонок (дата/телефон/ИНН/СНИЛС/...).
 *   R5 findAnomalies       — выбросы, дубли, пустоты.
 *   R6 getFormula          — формулы ячейки/диапазона.
 */

import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";

// ============================================================================
// R1. getWorkbookOverview — ОДИН context.sync() на всю книгу
// ============================================================================

interface SheetOverview {
  name: string;
  position: number;
  visible: boolean;
  rowCount: number;
  colCount: number;
  usedRangeAddress: string;
  headers: string[];
  /** Первые 3 строки данных (для контекста LLM). */
  sampleRows: unknown[][];
}

interface WorkbookOverview {
  sheets: SheetOverview[];
  tables: { name: string; sheetName: string; rangeAddress: string }[];
  activeSheetName: string;
  totalSheets: number;
  totalUsedCells: number;
}

/**
 * ФИКС N+1 (критично, см. docs/03-TOOLS-SPEC.md §2 R1):
 * До фикса — до 4 context.sync() на каждый лист (usedRange, headers, sample,
 * tables/pivots/charts) = ~80-100 round-trips на книге с 20 листами.
 *
 * Новый алгоритм:
 *   1. load worksheets.items + каждый sheet.usedRange + каждый headers/sample
 *      в одном общем batch БЕЗ sync.
 *   2. load tables/pivots/charts на уровне workbook.
 *   3. ОДИН context.sync() выполняет все накопленные запросы.
 *   4. Сборка overview в памяти.
 *
 * Цель: на книге с 20 листами 1 sync вместо 100. UX не зависает.
 */
export const getWorkbookOverviewTool = defineTool({
  name: "getWorkbookOverview",
  description: `Возвращает обзор всей книги: листы (имя, позиция, размеры, заголовки, образец 3 строк), таблицы и активный лист.
Используй ПЕРВЫМ при работе с незнакомой книгой, чтобы понять её структуру.
Не читает все данные — только метаданные и первые 3 строки каждого листа (экономит токены).`,
  parameters: {
    type: "object",
    properties: {},
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(_args: Record<string, unknown>): Promise<ToolResult> {
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      // Загружаем базовую информацию о листах одним вызовом
      worksheets.load("items/name, items/position, items/visibility");
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");

      // ОДИН sync для базовой информации
      await context.sync();

      const sheetCount = worksheets.items.length;
      const sheets: SheetOverview[] = [];
      let totalUsedCells = 0;

      // Если листов слишком много (>50) — ограничиваем для производительности
      const sheetsToInspect = worksheets.items.slice(0, 50);

      // Накапливаем load-запросы БЕЗ sync
      const usedRanges: Excel.Range[] = [];
      const headerRanges: Excel.Range[] = [];
      const sampleRanges: Excel.Range[] = [];

      for (const sheet of sheetsToInspect) {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("address, rowCount, columnCount, isNullObject");
        usedRanges.push(usedRange);
      }

      // ОДИН sync для всех usedRange сразу
      await context.sync();

      // Теперь знаем размеры каждого листа — накапливаем headers и sample
      for (let i = 0; i < sheetsToInspect.length; i++) {
        const sheet = sheetsToInspect[i];
        const usedRange = usedRanges[i];

        let rowCount = 0;
        let colCount = 0;
        let usedRangeAddress = "";
        let headers: string[] = [];
        let sampleRows: unknown[][] = [];

        if (!usedRange.isNullObject) {
          rowCount = usedRange.rowCount;
          colCount = usedRange.columnCount;
          usedRangeAddress = usedRange.address;
          totalUsedCells += rowCount * colCount;

          if (rowCount > 0 && colCount > 0) {
            // Headers: первая строка usedRange
            const headerRange = usedRange.getRow(0);
            headerRange.load("values");
            headerRanges.push(headerRange);

            // Sample: первые 3 строки (или меньше)
            const sampleRowCount = Math.min(3, rowCount);
            // Используем getResizedRange от первой строки
            const sampleStart = usedRange.getCell(0, 0);
            const sampleRange = sampleStart.getResizedRange(sampleRowCount - 1, colCount - 1);
            sampleRange.load("values");
            sampleRanges.push(sampleRange);
          }
        }

        sheets.push({
          name: sheet.name,
          position: sheet.position,
          visible: sheet.visibility === "Visible",
          rowCount,
          colCount,
          usedRangeAddress,
          headers, // заполним после sync
          sampleRows, // заполним после sync
        });
      }

      // ОДИН sync для всех headers + samples
      await context.sync();

      // Распределяем загруженные значения
      let headerIdx = 0;
      let sampleIdx = 0;
      for (let i = 0; i < sheets.length; i++) {
        const s = sheets[i];
        if (s.rowCount > 0 && s.colCount > 0) {
          if (headerIdx < headerRanges.length) {
            s.headers = (headerRanges[headerIdx].values[0] ?? []).map((v) =>
              String(v ?? ""),
            );
            headerIdx++;
          }
          if (sampleIdx < sampleRanges.length) {
            s.sampleRows = sampleRanges[sampleIdx].values as unknown[][];
            sampleIdx++;
          }
        }
      }

      // Таблицы на уровне workbook — одним запросом
      const tables = context.workbook.tables;
      tables.load("items/name, items/range/address, items/worksheet/name");
      await context.sync();

      const tablesOverview = tables.items.map((t) => ({
        name: t.name,
        sheetName: t.worksheet.name,
        rangeAddress: "",
      }));

      const overview: WorkbookOverview = {
        sheets,
        tables: tablesOverview,
        activeSheetName: activeSheet.name,
        totalSheets: sheetCount,
        totalUsedCells,
      };

      const summary = `Книга: ${sheetCount} лист(ов), ${tablesOverview.length} таблиц(ы), ~${totalUsedCells} ячеек. Активный лист: «${activeSheet.name}».`;

      return {
        ok: true,
        summary,
        data: overview,
      };
    });
  },
});

toolRegistry.registerDefinition(getWorkbookOverviewTool);

// ============================================================================
// R2. getRange — содержимое диапазона с типами
// ============================================================================

interface RangeCell {
  value: unknown;
  formula: string;
  numberFormat: string;
  type: "empty" | "string" | "number" | "boolean" | "date" | "error" | "formula";
}

interface RangeData {
  address: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: unknown[][];
  numberFormats: unknown[][];
  types: RangeCell["type"][][];
  headers: string[];
}

export function detectCellType(
  value: unknown,
  formula: string,
  numberFormat: string,
): RangeCell["type"] {
  if (formula && typeof formula === "string" && formula.startsWith("=")) {
    return "formula";
  }
  if (value === null || value === undefined || value === "") return "empty";
  if (typeof value === "number") {
    if (
      typeof numberFormat === "string" &&
      /[dy]/i.test(numberFormat) &&
      !/%/.test(numberFormat)
    ) {
      return "date";
    }
    return "number";
  }
  if (typeof value === "boolean") return "boolean";
  if (typeof value === "string") {
    if (/^#REF!|#VALUE!|#DIV\/0!|#NAME\?|#N\/A|#NULL!|#NUM!|#SPILL!|#CALC!/.test(value)) {
      return "error";
    }
    return "string";
  }
  return "string";
}

export const getRangeTool = defineTool({
  name: "getRange",
  description: `Возвращает содержимое диапазона: значения, формулы, форматы и определённые типы каждой ячейки.
Используй, чтобы прочитать конкретные данные перед вычислениями или изменениями.
Принимает адрес "A1", "A1:B10" или "Лист!A1:B10".`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес диапазона: "A1", "A1:B10" или "Лист1!A1:B10"',
      },
    },
    required: ["address"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: (_args: Record<string, unknown>) => 100,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес диапазона",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("address, values, formulas, numberFormat, rowCount, columnCount");
      await context.sync();

      const rowCount = range.rowCount;
      const columnCount = range.columnCount;
      const values = range.values as unknown[][];
      const formulas = range.formulas as unknown[][];
      const numberFormats = range.numberFormat as unknown[][];

      const types: RangeCell["type"][][] = [];
      for (let r = 0; r < values.length; r++) {
        const row: RangeCell["type"][] = [];
        for (let c = 0; c < (values[r]?.length ?? 0); c++) {
          row.push(
            detectCellType(
              values[r]?.[c],
              String(formulas[r]?.[c] ?? ""),
              String(numberFormats[r]?.[c] ?? ""),
            ),
          );
        }
        types.push(row);
      }

      const headers =
        rowCount > 0
          ? (values[0] ?? []).map((v) => String(v ?? ""))
          : [];

      const data: RangeData = {
        address: range.address,
        rowCount,
        columnCount,
        values,
        formulas,
        numberFormats,
        types,
        headers,
      };

      const summary = `Прочитан диапазон ${range.address}: ${rowCount} строк × ${columnCount} колонк(а/ок).`;

      return { ok: true, summary, data };
    });
  },
});

toolRegistry.registerDefinition(getRangeTool);

// ============================================================================
// R6. getFormula — формулы ячейки/диапазона (без значений)
// ============================================================================

export const getFormulaTool = defineTool({
  name: "getFormula",
  description: `Возвращает ТОЛЬКО формулы диапазона (без значений) — полезно для отладки или когда нужно понять структуру вычислений.
Если ячейка содержит не формулу — возвращает пустую строку.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес ячейки или диапазона: "C5" или "C5:F10"',
      },
    },
    required: ["address"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 1,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("formulas, address, rowCount, columnCount");
      await context.sync();

      const formulas = range.formulas as string[][];
      let formulaCount = 0;
      const flat: string[] = [];
      for (const row of formulas) {
        for (const f of row) {
          if (typeof f === "string" && f.startsWith("=")) formulaCount++;
          flat.push(f);
        }
      }

      const summary = `${range.address}: ${formulaCount} формул(а/ы) из ${flat.length} ячеек.`;

      return {
        ok: true,
        summary,
        data: {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          formulas,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(getFormulaTool);

// ============================================================================
// R3. getRangeStats — вычисления min/max/sum/avg в JS
// ============================================================================

interface ColumnStats {
  index: number;
  header?: string;
  numericCount: number;
  nonEmptyCount: number;
  emptyCount: number;
  uniqueCount: number;
  sum: number | null;
  min: number | null;
  max: number | null;
  avg: number | null;
  sampleValues: unknown[];
}

interface RangeStatsResult {
  address: string;
  rowCount: number;
  columnCount: number;
  columns: ColumnStats[];
}

/**
 * Парсит "числовую" строку в число: "1 234,5" → 1234.5, "1,234.5" → 1234.5.
 * Возвращает null, если значение не похоже на число.
 */
export function parseNumeric(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  let normalized: string;
  if (trimmed.includes(".") && trimmed.includes(",")) {
    normalized = trimmed.replace(/,/g, "");
  } else if (trimmed.includes(",")) {
    normalized = trimmed.replace(/\s/g, "").replace(",", ".");
  } else {
    normalized = trimmed.replace(/\s/g, "");
  }

  const percentMatch = normalized.match(/^-?\d+(?:\.\d+)?%$/);
  if (percentMatch) {
    return parseFloat(normalized.replace("%", "")) / 100;
  }

  const moneyMatch = normalized.match(/^-?[\d.]+$/);
  if (moneyMatch) {
    const n = parseFloat(normalized);
    return Number.isFinite(n) ? n : null;
  }

  return null;
}

export const getRangeStatsTool = defineTool({
  name: "getRangeStats",
  description: `Считает статистику по диапазону в JS: min, max, sum, avg, count, unique, empty — для каждой колонки отдельно.
Используй ПЕРЕД формулами типа SUM/AVG, чтобы понять данные. Распознаёт числа-как-текст ("1 234,5" → 1234.5) и проценты.
Первая строка трактуется как заголовки колонок.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес диапазона с данными, напр. "A1:D100"',
      },
    },
    required: ["address"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 100,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("values, address, rowCount, columnCount");
      await context.sync();

      const values = range.values as unknown[][];
      const rowCount = range.rowCount;
      const columnCount = range.columnCount;

      if (rowCount === 0 || columnCount === 0) {
        return {
          ok: true,
          summary: `Диапазон ${range.address} пуст`,
          data: { address: range.address, rowCount, columnCount, columns: [] },
        };
      }

      const headers = (values[0] ?? []).map((v) => String(v ?? ""));
      const dataRows = values.slice(1);

      const columns: ColumnStats[] = [];
      for (let c = 0; c < columnCount; c++) {
        const colValues: unknown[] = [];
        const numericValues: number[] = [];
        const uniqueSet = new Set<string>();
        let emptyCount = 0;

        for (const row of dataRows) {
          const v = row?.[c];
          colValues.push(v);

          if (v === null || v === undefined || v === "") {
            emptyCount++;
            continue;
          }

          uniqueSet.add(String(v));
          const num = parseNumeric(v);
          if (num !== null) numericValues.push(num);
        }

        const nonEmptyCount = colValues.length - emptyCount;
        const sum = numericValues.length > 0
          ? numericValues.reduce((a, b) => a + b, 0)
          : null;
        const min = numericValues.length > 0
          ? Math.min(...numericValues)
          : null;
        const max = numericValues.length > 0
          ? Math.max(...numericValues)
          : null;
        const avg = numericValues.length > 0 && sum !== null
          ? sum / numericValues.length
          : null;

        const samples: unknown[] = [];
        for (const v of colValues) {
          if (v !== null && v !== undefined && v !== "") {
            samples.push(v);
            if (samples.length >= 3) break;
          }
        }

        columns.push({
          index: c,
          header: headers[c],
          numericCount: numericValues.length,
          nonEmptyCount,
          emptyCount,
          uniqueCount: uniqueSet.size,
          sum,
          min,
          max,
          avg,
          sampleValues: samples,
        });
      }

      const data: RangeStatsResult = {
        address: range.address,
        rowCount,
        columnCount,
        columns,
      };

      const totalNumeric = columns.reduce((a, col) => a + col.numericCount, 0);
      const summary = `Статистика по ${range.address}: ${rowCount - 1} строк данных, ${columnCount} колонк(а/ок), ${totalNumeric} числовых значений.`;

      return { ok: true, summary, data };
    });
  },
});

toolRegistry.registerDefinition(getRangeStatsTool);

// ============================================================================
// R4. detectDataTypes — типы колонок (дата/телефон/ИНН/СНИЛС/email/...)
// ============================================================================

export type ColumnType =
  | "integer"
  | "float"
  | "currency"
  | "percent"
  | "date"
  | "datetime"
  | "time"
  | "phone"
  | "email"
  | "url"
  | "inn"
  | "kpp"
  | "ogrn"
  | "snils"
  | "text"
  | "mixed"
  | "empty";

export interface ColumnTypeResult {
  index: number;
  header?: string;
  type: ColumnType;
  /** Доля значений, подходящих под type (0..1). */
  confidence: number;
  /** 3 примера значений. */
  sampleValues: unknown[];
}

export interface DetectDataTypesResult {
  address: string;
  rowCount: number;
  columnCount: number;
  columns: ColumnTypeResult[];
}

// ---------------------------------------------------------------------------
// Детекторы типов — чистые функции для тестирования
// ---------------------------------------------------------------------------

/** ИНН физлица (12 цифр) или юрлица (10 цифр). */
export function isInn(value: unknown): boolean {
  if (typeof value !== "string" && typeof value !== "number") return false;
  const s = String(value).replace(/\D/g, "");
  return s.length === 10 || s.length === 12;
}

/** КПП — 9 цифр. */
export function isKpp(value: unknown): boolean {
  if (typeof value !== "string" && typeof value !== "number") return false;
  const s = String(value).replace(/\D/g, "");
  return s.length === 9;
}

/** ОГРН (13) или ОГРНИП (15). */
export function isOgrn(value: unknown): boolean {
  if (typeof value !== "string" && typeof value !== "number") return false;
  const s = String(value).replace(/\D/g, "");
  return s.length === 13 || s.length === 15;
}

/** СНИЛС: "123-456-789 00" или "12345678900" (11 цифр). */
export function isSnils(value: unknown): boolean {
  if (typeof value !== "string" && typeof value !== "number") return false;
  const s = String(value).replace(/\D/g, "");
  return s.length === 11;
}

/** Российский телефон: +7/8 + 10 цифр, либо 10 цифр с телефонными разделителями. */
export function isPhoneRu(value: unknown): boolean {
  if (typeof value !== "string" && typeof value !== "number") return false;
  const raw = String(value).trim();
  if (!raw) return false;
  const s = raw.replace(/\D/g, "");
  // 11 цифр, начинается с 7/8 — классический мобильный RU (с кодом страны).
  if (s.length === 11 && (s.startsWith("7") || s.startsWith("8"))) return true;
  // 10 цифр — только при наличии телефонных разделителей (+, скобки, дефисы, пробелы),
  // иначе ИНН юрлица (10 цифр) и прочие «голые» числа ложно детектируются как телефон.
  if (s.length === 10 && /[()+\-\s]/.test(raw)) return true;
  return false;
}

/** Email. */
export function isEmail(value: unknown): boolean {
  if (typeof value !== "string") return false;
  return /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(value.trim());
}

/** URL. */
export function isUrl(value: unknown): boolean {
  if (typeof value !== "string") return false;
  return /^https?:\/\/[^\s]+$/i.test(value.trim());
}

/** Дата: DD.MM.YYYY, DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD и варианты с временем. */
export function isDateString(value: unknown): boolean {
  if (value === null || value === undefined) return false;
  // Excel serial date number (только если value — number и в правдоподобном диапазоне)
  if (typeof value === "number") {
    // 1 = 1900-01-01, 60000 ≈ 2064 год — правдоподобный диапазон дат
    return value >= 1 && value <= 90000;
  }
  if (typeof value !== "string") return false;
  const s = value.trim();
  // DD.MM.YYYY или DD.MM.YY
  if (/^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}$/.test(s)) return true;
  // YYYY-MM-DD (ISO)
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) return true;
  // DD.MM.YYYY HH:MM
  if (/^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\s+\d{1,2}:\d{2}(:\d{2})?$/.test(s)) return true;
  return false;
}

/** Определяет тип одного значения. */
export function detectValueType(value: unknown): ColumnType {
  if (value === null || value === undefined || value === "") return "empty";

  if (typeof value === "number") {
    // Длинные числа могут быть ИНН/ОГРН, записанными без текста.
    if (isInn(value)) return "inn";
    if (isOgrn(value)) return "ogrn";
    // Дата по одному числу не определяется — нужен numberFormat ячейки,
    // которого у detectValueType нет. Поэтому число = integer/float.
    return Number.isInteger(value) ? "integer" : "float";
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return "empty";

    // Приоритет: специфичные RU-форматы ПЕРВЕЕ общих чисел.
    // Телефон проверяем до СНИЛС: оба дают 11 цифр, но телефон имеет более
    // узкий маркер (код страны 7/8 или телефонные разделители).
    if (isPhoneRu(value)) return "phone";
    if (isInn(value)) return "inn";
    if (isSnils(value)) return "snils";
    if (isKpp(value)) return "kpp";
    if (isOgrn(value)) return "ogrn";
    if (isEmail(value)) return "email";
    if (isUrl(value)) return "url";
    if (isDateString(value)) return "date";

    // Процент
    if (/^-?\d+(?:[.,]\d+)?%$/.test(trimmed)) return "percent";
    // Валюта: "1 234,5 ₽", "$100", "100 руб."
    if (/^[$€₽]|руб\.?$/i.test(trimmed)) return "currency";
    // Число как текст
    const num = parseNumeric(trimmed);
    if (num !== null) {
      return Number.isInteger(num) ? "integer" : "float";
    }
    return "text";
  }
  return "text";
}

/**
 * Определяет доминирующий тип колонки по массиву значений.
 * Возвращает { type, confidence } — confidence это доля значений, подходящих
 * под type (исключая empty).
 */
export function detectColumnType(
  values: unknown[],
): { type: ColumnType; confidence: number } {
  const nonEmpty = values.filter(
    (v) => v !== null && v !== undefined && v !== "",
  );
  if (nonEmpty.length === 0) {
    return { type: "empty", confidence: 1 };
  }

  const counts = new Map<ColumnType, number>();
  for (const v of nonEmpty) {
    const t = detectValueType(v);
    counts.set(t, (counts.get(t) ?? 0) + 1);
  }

  // Находим тип с максимальным count
  let bestType: ColumnType = "text";
  let bestCount = 0;
  for (const [t, c] of counts) {
    if (c > bestCount) {
      bestCount = c;
      bestType = t;
    }
  }

  const confidence = bestCount / nonEmpty.length;
  // Если confidence < 0.6 — колонка mixed
  const type: ColumnType = confidence < 0.6 ? "mixed" : bestType;

  return { type, confidence };
}

// ---------------------------------------------------------------------------
// Инструмент detectDataTypes
// ---------------------------------------------------------------------------

export const detectDataTypesTool = defineTool({
  name: "detectDataTypes",
  description: `Определяет реальные типы данных каждой колонки: дата, число, телефон, email, ИНН, СНИЛС, валюта, процент, текст.
Используй ПЕРЕД очисткой/форматированием, чтобы понять с чем имеешь дело.
Распознаёт российские форматы: ИНН (10/12 цифр), СНИЛС, телефон (+7/8), дата ДД.ММ.ГГГГ.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес диапазона с заголовком в первой строке, напр. "A1:D100"',
      },
    },
    required: ["address"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 100,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("values, address, rowCount, columnCount");
      await context.sync();

      const values = range.values as unknown[][];
      const rowCount = range.rowCount;
      const columnCount = range.columnCount;

      if (rowCount <= 1 || columnCount === 0) {
        return {
          ok: true,
          summary: `Недостаточно данных в ${range.address}`,
          data: { address: range.address, rowCount, columnCount, columns: [] },
        };
      }

      const headers = (values[0] ?? []).map((v) => String(v ?? ""));
      const dataRows = values.slice(1);

      const columns: ColumnTypeResult[] = [];
      for (let c = 0; c < columnCount; c++) {
        const colValues = dataRows.map((row) => row?.[c]);
        const { type, confidence } = detectColumnType(colValues);

        const samples: unknown[] = [];
        for (const v of colValues) {
          if (v !== null && v !== undefined && v !== "") {
            samples.push(v);
            if (samples.length >= 3) break;
          }
        }

        columns.push({
          index: c,
          header: headers[c],
          type,
          confidence,
          sampleValues: samples,
        });
      }

      const data: DetectDataTypesResult = {
        address: range.address,
        rowCount,
        columnCount,
        columns,
      };

      const summary = `Типы колонок в ${range.address}: ${columns
        .map((c) => `${c.header ?? `#${c.index}`}=${c.type}`)
        .join(", ")}.`;

      return { ok: true, summary, data };
    });
  },
});

toolRegistry.registerDefinition(detectDataTypesTool);

// ============================================================================
// R5. findAnomalies — выбросы, дубли, пустоты, неразборные, ошибки Excel
// (docs/03-TOOLS-SPEC.md §1 R5, эталонный сценарий №19 «Объяснить что в таблице»)
// ============================================================================

/**
 * Строка-аналия в колонке. row=1 — первая строка данных (после заголовка).
 * kind определяет категорию проблемы.
 */
export type AnomalyKind =
  | "empty" // пусто там, где в колонке есть значения
  | "duplicate" // повтор значения в колонке-признаке (ID/email/ИНН/…)
  | "outlier" // числовой выброс (Z-score или IQR)
  | "type_mismatch" // значение не подходит под доминирующий тип колонки
  | "error"; // значение ошибки Excel (#Н/Д, #ДЕЛ/0!, …)

export interface Anomaly {
  row: number;
  column: number;
  header?: string;
  kind: AnomalyKind;
  value: unknown;
  /** Человекочитаемое пояснение для LLM/UI. */
  detail: string;
}

export interface ColumnAnomalyReport {
  index: number;
  header?: string;
  anomalies: Anomaly[];
}

export interface FindAnomaliesResult {
  address: string;
  rowCount: number;
  columnCount: number;
  /** Общее число найденных аномалий (для summary/прогресса). */
  totalAnomalies: number;
  columns: ColumnAnomalyReport[];
}

// ---------------------------------------------------------------------------
// Детекторы — чистые функции для unit-тестов
// ---------------------------------------------------------------------------

/** Регэксп значений-ошибок Excel (RU и EN локали). */
const ERROR_RE = /^#(REF|VALUE|DIV\/0|NAME\?|N\/A|NULL!|NUM!|SPILL!|CALC!|ДЕЛ\/0!|ЗНАЧ!|ИМЯ\?|Н\/Д|ЧИСЛО!|ПУСТО!)/i;

/** true для значения-ошибки Excel ("#Н/Д", "#REF!", "#VALUE!" и пр.). */
export function isErrorValue(value: unknown): boolean {
  if (typeof value !== "string") return false;
  return ERROR_RE.test(value.trim());
}

/**
 * Числовой выброс через МОДИФИЦИРОВАННЫЙ Z-score (робастный к маскированию).
 *
 * Классический Z-score (|x−mean|/std) ломается, когда один гигантский выброс
 * раздувает mean и std так, что сам перестаёт считаться выбросом. Поэтому
 * используем median + MAD (median absolute deviation) — они не чувствительны
 * к крайним значениям (Iglewicz & Hoaglin, NIST).
 *
 * scale = MAD × 1.4826 (оценка std по MAD). Если MAD=0 (большинство значений
 * совпадают с медианой) — fallback на mean absolute deviation × 1.2533.
 *
 * Применяется при ≥8 чисел. threshold 3.5 — стандартная рекомендация NIST.
 */
export function detectOutliersZScore(
  nums: number[],
  threshold = 3.5,
): number[] {
  if (nums.length < 8) return [];

  const sorted = [...nums].sort((a, b) => a - b);
  const n = sorted.length;
  const mid = Math.floor(n / 2);
  const median =
    n % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];

  const absDevs = nums.map((x) => Math.abs(x - median));
  const sortedDevs = [...absDevs].sort((a, b) => a - b);
  const mad =
    n % 2 === 0
      ? (sortedDevs[mid - 1] + sortedDevs[mid]) / 2
      : sortedDevs[mid];

  // scale — оценка std робастным методом.
  let scale: number;
  if (mad > 0) {
    scale = mad * 1.4826;
  } else {
    // MAD=0: большинство значений = медиана. Fallback на mean abs deviation
    // (×1.2533 — поправка, т.к. mean abs dev ≈ std × 0.7979).
    const meanAbsDev = absDevs.reduce((a, b) => a + b, 0) / n;
    if (meanAbsDev === 0) return []; // все значения одинаковы
    scale = meanAbsDev * 1.2533;
  }

  const outliers: number[] = [];
  for (let i = 0; i < nums.length; i++) {
    if (Math.abs(nums[i] - median) / scale > threshold) outliers.push(i);
  }
  return outliers;
}

/**
 * Числовой выброс через IQR (межквартильный размах).
 * Используется для малых выборок (<8 чисел) и как замена Z-score при std=0.
 * Значение — выброс, если x < Q1 − 1.5·IQR или x > Q3 + 1.5·IQR.
 */
export function detectOutliersIQR(nums: number[]): number[] {
  if (nums.length < 4) return [];
  const sorted = [...nums].sort((a, b) => a - b);
  const q1 = sorted[Math.floor(sorted.length * 0.25)];
  const q3 = sorted[Math.floor(sorted.length * 0.75)];
  const iqr = q3 - q1;
  if (iqr === 0) return [];
  const lower = q1 - 1.5 * iqr;
  const upper = q3 + 1.5 * iqr;
  const outliers: number[] = [];
  for (let i = 0; i < nums.length; i++) {
    if (nums[i] < lower || nums[i] > upper) outliers.push(i);
  }
  return outliers;
}

/**
 * Колонка считается «колонкой-признаком» (где дубли подозрительны), если её
 * доминирующий тип — inn/snils/kpp/ogrn/email/phone (ID-подобные значения).
 * Для произвольного text/integer дубли НЕ flagged — повторяющиеся значения
 * там нормальны (количество, категория, комментарий). Эвристика по uniqueRatio
 * убрана как дающая ложные срабатывания (P6).
 */
function isLikelyKeyColumn(dominantType: ColumnType): boolean {
  const idTypes: ColumnType[] = [
    "inn",
    "snils",
    "kpp",
    "ogrn",
    "email",
    "phone",
  ];
  return idTypes.includes(dominantType);
}

/**
 * Находит аномалии в одной колонке.
 * @param colValues значения колонки БЕЗ заголовка (только данные).
 * @param header     имя колонки.
 * @param index      индекс колонки.
 */
export function findColumnAnomalies(
  colValues: unknown[],
  header: string | undefined,
  index: number,
): Anomaly[] {
  const anomalies: Anomaly[] = [];
  const total = colValues.length;

  // 1. Пустоты и ошибки — по каждой ячейке.
  const nonEmptyValues: { value: unknown; row: number }[] = [];
  const numericValues: { num: number; row: number }[] = [];

  for (let i = 0; i < total; i++) {
    const v = colValues[i];
    const row = i + 1; // row=1 — первая строка данных
    if (v === null || v === undefined || v === "") {
      anomalies.push({
        row,
        column: index,
        header,
        kind: "empty",
        value: v,
        detail: "Пустая ячейка в колонке с данными",
      });
      continue;
    }
    if (isErrorValue(v)) {
      anomalies.push({
        row,
        column: index,
        header,
        kind: "error",
        value: v,
        detail: `Значение-ошибка Excel: ${v}`,
      });
      continue;
    }
    nonEmptyValues.push({ value: v, row });
    const num = parseNumeric(v);
    if (num !== null) numericValues.push({ num, row });
  }

  // 2. Дубли — только в колонках-признаках (ID/email/…), не для произвольного текста.
  const { type: dominantType } = detectColumnType(colValues);
  if (isLikelyKeyColumn(dominantType)) {
    const seen = new Map<string, number>(); // значение → первая встреченная строка
    for (const { value, row } of nonEmptyValues) {
      const key = String(value);
      const firstRow = seen.get(key);
      if (firstRow !== undefined) {
        anomalies.push({
          row,
          column: index,
          header,
          kind: "duplicate",
          value,
          detail: `Дубликат значения из строки ${firstRow}`,
        });
      } else {
        seen.set(key, row);
      }
    }
  }

  // 3. Числовые выбросы — Z-score (≥8 чисел), иначе IQR.
  if (numericValues.length >= 4) {
    const nums = numericValues.map((x) => x.num);
    const outlierIdxs =
      nums.length >= 8 ? detectOutliersZScore(nums) : detectOutliersIQR(nums);
    const seenOutlierRows = new Set<number>();
    for (const idx of outlierIdxs) {
      const { num, row } = numericValues[idx];
      // IQR может вернуть одно и то же число несколько раз при равенствах —
      // страховка от дублирования аномалии на одну строку.
      if (seenOutlierRows.has(row)) continue;
      seenOutlierRows.add(row);
      anomalies.push({
        row,
        column: index,
        header,
        kind: "outlier",
        value: num,
        detail: `Числовой выброс: ${num}`,
      });
    }
  }

  // 4. Несоответствие типу — значение не подходит под доминирующий тип колонки.
  if (dominantType !== "mixed" && dominantType !== "empty" && dominantType !== "text") {
    for (const { value, row } of nonEmptyValues) {
      const t = detectValueType(value);
      if (t !== dominantType && t !== "empty") {
        anomalies.push({
          row,
          column: index,
          header,
          kind: "type_mismatch",
          value,
          detail: `Ожидается «${dominantType}», но похоже на «${t}»`,
        });
      }
    }
  }

  return anomalies;
}

// ---------------------------------------------------------------------------
// Инструмент findAnomalies
// ---------------------------------------------------------------------------

export const findAnomaliesTool = defineTool({
  name: "findAnomalies",
  description: `Находит аномалии в данных диапазона: пустые ячейки, дубликаты в колонках-ID (email/ИНН/СНИЛС/телефон), числовые выбросы (Z-score/IQR), значения-ошибки Excel (#Н/Д, #ДЕЛ/0!) и значения не подходящего типа.
Используй ПЕРЕД очисткой, чтобы понять где проблемы, или когда пользователь просит «найди ошибки/выбросы/дубли».
Первая строка трактуется как заголовки.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес диапазона с заголовком в первой строке, напр. "A1:D100"',
      },
    },
    required: ["address"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 100,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("values, address, rowCount, columnCount");
      await context.sync();

      const values = range.values as unknown[][];
      const rowCount = range.rowCount;
      const columnCount = range.columnCount;

      if (rowCount <= 1 || columnCount === 0) {
        return {
          ok: true,
          summary: `Недостаточно данных в ${range.address}`,
          data: {
            address: range.address,
            rowCount,
            columnCount,
            totalAnomalies: 0,
            columns: [],
          } as FindAnomaliesResult,
        };
      }

      const headers = (values[0] ?? []).map((v) => String(v ?? ""));
      const dataRows = values.slice(1);

      const columns: ColumnAnomalyReport[] = [];
      let totalAnomalies = 0;
      for (let c = 0; c < columnCount; c++) {
        const colValues = dataRows.map((row) => row?.[c]);
        const anomalies = findColumnAnomalies(colValues, headers[c], c);
        totalAnomalies += anomalies.length;
        columns.push({ index: c, header: headers[c], anomalies });
      }

      const data: FindAnomaliesResult = {
        address: range.address,
        rowCount,
        columnCount,
        totalAnomalies,
        columns,
      };

      // Краткая разбивка по видам аномалий для summary.
      const byKind = new Map<AnomalyKind, number>();
      for (const col of columns) {
        for (const a of col.anomalies) {
          byKind.set(a.kind, (byKind.get(a.kind) ?? 0) + 1);
        }
      }
      const kindLabels: Record<AnomalyKind, string> = {
        empty: "пустых",
        duplicate: "дублей",
        outlier: "выбросов",
        type_mismatch: "несоответствий типа",
        error: "ошибок Excel",
      };
      const breakdown = Array.from(byKind.entries())
        .map(([k, n]) => `${n} ${kindLabels[k]}`)
        .join(", ");

      const summary =
        totalAnomalies === 0
          ? `Аномалий не найдено в ${range.address} (${rowCount - 1} строк × ${columnCount} колонк(а/ок)).`
          : `Найдено ${totalAnomalies} аномалий в ${range.address}: ${breakdown}.`;

      return { ok: true, summary, data, cellsAffected: totalAnomalies };
    });
  },
});

toolRegistry.registerDefinition(findAnomaliesTool);
