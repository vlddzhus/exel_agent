/**
 * tools/index.ts — barrel-файл для side-effect регистрации всех инструментов.
 *
 * Каждый модуль инструментов регистрирует себя через top-level вызов
 * `toolRegistry.registerDefinition(...)` (новый API) или `toolRegistry.register(...)`
 * (legacy). Пока модуль не импортирован, реестр остаётся ПУСТЫМ.
 *
 * До этого файла ни один React-модуль не импортировал инструменты → useAgent
 * отправлял в бэкенд пустой массив tools → у LLM не было инструментов для вызова
 * (корень разрыва цепи «Мозг ↔ Руки», см. отчёт по гибридному циклу).
 *
 * Использование: один side-effect import в точке входа (index.tsx) регистрирует
 * все инструменты до рендера App:
 *   import "./tools";
 *
 * Порядок не важен — имена уникальны (registry бросает при дубликате).
 */

// Read: R1-R6
import "./read";
// Write: W1-W5
import "./write";
// Format: F1-F5
import "./format";
// Format-extended: F7-F11 (setRowHeights, autoFitRows, copyFormat, applyNamedStyle, setSheetTabColor)
import "./format-extended";
// AutoDesigner: applyAutoDesign — авто-оформление таблиц по типу данных
import "./auto-designer";
// Structure: S1-S5
import "./structure";
// Transform: T1-T6
import "./transform";

// Format: F6 — setColumnWidths (добавлен в format.ts)
// Validation: V1 — addDataValidation
import "./data-validation";

// Advanced Structure: Tier 1 — manageRowsColumns, mergeCells, unmergeCells, manageSheetProtection
import "./advanced-structure";

// ── Итерация «Расширение инструментов» (HIGH-9) ──
// Search: findAndReplace — поиск и замена по диапазону/листу
import "./search-tools";
// Named Ranges: manageNamedRanges — именованные диапазоны (add/list/delete/get)
import "./named-ranges";
// Hyperlinks: manageHyperlinks — гиперссылки в ячейках (add/get/remove)
import "./hyperlinks";
// Grouping: manageGrouping — группировка строк/столбцов (сворачиваемые секции)
import "./grouping";
// Sheet View: manageSheetView (сетка/масштаб), managePageSetup (печать)
import "./sheet-view";
// Comments & Charts: manageComments, formatChart (пост-форматирование графиков)
import "./comments-charts";

// Excel Table-инструменты (legacy register API): listTables, createTable,
// addTableRow, sortTable, filterTable — уникальные операции над Table-объектами,
// не перекрытые canonical-инструментами.
import "./table-tools";

// ВНИМАНИЕ: formula-tools.ts намеренно НЕ импортируется здесь. Его инструменты
// (setFormula, applyFormat, setCellFormat, fillFormula) — legacy-дубликаты
// canonical из write.ts/format.ts и порождают конфликт имён ("Tool already
// registered: setFormula"). Мёртвый код — удалить полностью в отдельном коммите
// (AGENTS.md правило 6); пока исключён из barrel, чтобы React-надстройка
// стартовала без падения регистрации.

export { toolRegistry } from "./registry";
export type {
  ToolDefinition,
  ToolResult,
  ToolContext,
  RiskLevel,
} from "./registry";
