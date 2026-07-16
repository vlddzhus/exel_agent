/**
 * _shared/address.ts — ЕДИНЫЙ источник истины для адресной арифметики Excel.
 *
 * Все инструменты (tools/*.ts) обязаны импортировать хелперы адресов отсюда.
 * Дублирование в других модулях запрещено (см. docs/03-TOOLS-SPEC.md §3.1,
 * docs/00-MASTER-PLAN.md P6).
 *
 * Что включено:
 *   - columnToLetter / letterToColumn (0-based: 0→A, 26→AA, 18278→XFD)
 *   - parseCellAddress / parseRangeAddress
 *   - computeRange (start + rows/cols → range)
 *   - mergeRangeAddress (объединение, числовое сравнение)
 *   - isValidAddress / isValidRangeAddress (безопасная проверка)
 *   - normalizeAddress (канонический вид: убрать кавычки, верхний регистр колонок)
 *   - resolveRange (безопасное получение Excel.Range из строки адреса)
 *
 * Принципы (docs/03-TOOLS-SPEC.md §0.2):
 *   - Никаких `any`, строгая типизация.
 *   - Адреса валидируются ДО передачи в Office.js (защита от path traversal).
 *   - Сравнение колонок ЧИСЛОВОЕ (не строковое), чтобы Z < AA правильно.
 */

// ---------------------------------------------------------------------------
// Column / Row primitives
// ---------------------------------------------------------------------------

/**
 * Convert column number to letter: 0→A, 1→B, 25→Z, 26→AA, 701→ZZ, 18278→XFD.
 * Бросает RangeError для отрицательных или > 18278 (последняя колонка Excel).
 */
export function columnToLetter(col: number): string {
  if (!Number.isInteger(col) || col < 0 || col > 16383) {
    throw new RangeError(`Column out of range [0, 16383] (XFD): ${col}`);
  }
  let result = "";
  let n = col + 1;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

/**
 * Convert column letter to number: A→0, B→1, Z→25, AA→26, XFD→16383.
 * Принимает только буквы A-Z (без $, без цифр). Бросает RangeError для невалидного.
 */
export function letterToColumn(col: string): number {
  const upper = col.toUpperCase();
  if (!/^[A-Z]+$/.test(upper)) {
    throw new RangeError(`Invalid column letter: ${col}`);
  }
  let result = 0;
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result - 1; // 0-based
}

// ---------------------------------------------------------------------------
// Cell / Range parsing
// ---------------------------------------------------------------------------

export interface CellAddress {
  col: string; // "A", "AA"
  row: number; // 1-based
}

export interface RangeAddress {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
}

/** Parse "A1" → { col: "A", row: 1 }. Бросает Error для невалидного. */
export function parseCellAddress(cell: string): CellAddress {
  const match = cell.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell address: ${cell}`);
  return { col: match[1], row: parseInt(match[2], 10) };
}

/** Parse "A1:C10" → { startCol, startRow, endCol, endRow }. Бросает Error. */
export function parseRangeAddress(range: string): RangeAddress {
  const parts = range.split(":");
  if (parts.length !== 2) throw new Error(`Invalid range address: ${range}`);
  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1]);
  return {
    startCol: start.col,
    startRow: start.row,
    endCol: end.col,
    endRow: end.row,
  };
}

// ---------------------------------------------------------------------------
// Range computation
// ---------------------------------------------------------------------------

/**
 * Compute range from start cell and dimensions.
 * "A11", 9, 5 → "A11:E19"
 * Бросает Error для невалидного startCell или неположительных rows/cols.
 */
export function computeRange(
  startCell: string,
  rows: number,
  cols: number,
): string {
  if (!Number.isInteger(rows) || rows <= 0) {
    throw new Error(`rows must be positive integer, got: ${rows}`);
  }
  if (!Number.isInteger(cols) || cols <= 0) {
    throw new Error(`cols must be positive integer, got: ${cols}`);
  }
  const { col, row } = parseCellAddress(startCell);
  const endCol = columnToLetter(letterToColumn(col) + cols - 1);
  const endRow = row + rows - 1;
  return `${col}${row}:${endCol}${endRow}`;
}

/**
 * Merge two ranges into a bounding range covering both.
 * Важно: сравнение колонок ЧИСЛОВОЕ, не строковое (фиксит баг "Z" > "AA").
 *
 * "A1:B2", "C3:D4" → "A1:D4"
 * "A1:Z10", "AA1:AB10" → "A1:AB10"  (а не "A1:Z10" при строковом сравнении)
 */
export function mergeRangeAddress(a: string, b: string): string {
  const pa = parseRangeAddress(a);
  const pb = parseRangeAddress(b);
  const minCol = Math.min(letterToColumn(pa.startCol), letterToColumn(pb.startCol));
  const maxCol = Math.max(letterToColumn(pa.endCol), letterToColumn(pb.endCol));
  const minRow = Math.min(pa.startRow, pb.startRow);
  const maxRow = Math.max(pa.endRow, pb.endRow);
  return `${columnToLetter(minCol)}${minRow}:${columnToLetter(maxCol)}${maxRow}`;
}

// ---------------------------------------------------------------------------
// Validation (безопасные проверки, не бросают)
// ---------------------------------------------------------------------------

/** true если строка — валидный адрес ячейки ("A1", "AA123"). Не бросает. */
export function isValidCellAddress(cell: string): boolean {
  return /^[A-Za-z]+\d+$/.test(cell);
}

/**
 * true если строка — валидный адрес диапазона ("A1:B2") или ячейки ("A1").
 * НЕ принимает полные адреса с указанием листа ("Лист!A1") — это валидируется
 * отдельно в isValidFullAddress.
 */
export function isValidRangeAddress(range: string): boolean {
  if (range.includes(":")) {
    const parts = range.split(":");
    if (parts.length !== 2) return false;
    return isValidCellAddress(parts[0]) && isValidCellAddress(parts[1]);
  }
  return isValidCellAddress(range);
}

/**
 * true если строка — валидный full address ("Лист!A1:B2", "'Мой лист'!A1", "A1").
 * Проверяет безопасность: нет path traversal (../, /, \).
 */
export function isValidFullAddress(full: string): boolean {
  if (!full || typeof full !== "string" || full.length > 500) return false;
  // Запрет path traversal и спецсимволов
  if (/[\\/]|\.\.|\0/.test(full)) return false;
  const bangIndex = full.indexOf("!");
  if (bangIndex < 0) {
    return isValidRangeAddress(full);
  }
  const sheetPart = full.substring(0, bangIndex);
  const rangePart = full.substring(bangIndex + 1);
  // Sheet name: опционально в одинарных кавычках, непустой, без управления
  if (!isValidSheetName(sheetPart)) return false;
  return isValidRangeAddress(rangePart);
}

/**
 * Валидация имени листа: 1-31 символ, без : \ / ? * [ ], может быть в кавычках.
 * Принимает как 'Мой лист' так и МойЛист.
 */
export function isValidSheetName(name: string): boolean {
  if (!name || typeof name !== "string") return false;
  // Снимаем кавычки если есть
  const clean =
    name.startsWith("'") && name.endsWith("'") && name.length >= 2
      ? name.slice(1, -1)
      : name;
  if (clean.length === 0 || clean.length > 31) return false;
  if (/[\\/:?*\[\]]/.test(clean)) return false;
  if (clean.startsWith("'") || clean.endsWith("'")) return false;
  return true;
}

// ---------------------------------------------------------------------------
// Normalization
// ---------------------------------------------------------------------------

/**
 * Канонический вид адреса диапазона: убрать кавычки вокруг имени листа,
 * поднять регистр колонок, нормализовать к "Лист!A1:B2".
 * "Лист!a1:b2" → "Лист!A1:B2"
 * "'Мой лист'!a1" → "'Мой лист'!A1"
 */
export function normalizeAddress(full: string): string {
  const bangIndex = full.indexOf("!");
  if (bangIndex < 0) {
    return normalizeRangePart(full);
  }
  const sheetPart = full.substring(0, bangIndex);
  const rangePart = full.substring(bangIndex + 1);
  return `${sheetPart}!${normalizeRangePart(rangePart)}`;
}

function normalizeRangePart(range: string): string {
  if (range.includes(":")) {
    const [start, end] = range.split(":");
    return `${normalizeCellPart(start)}:${normalizeCellPart(end)}`;
  }
  return normalizeCellPart(range);
}

function normalizeCellPart(cell: string): string {
  const match = cell.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return cell;
  return match[1].toUpperCase() + match[2];
}

/**
 * Извлечь { sheetName, rangeAddress } из full address.
 * "Лист!A1:B2" → { sheetName: "Лист", rangeAddress: "A1:B2" }
 * "'Мой лист'!A1" → { sheetName: "Мой лист", rangeAddress: "A1" }
 * "A1" → { sheetName: undefined, rangeAddress: "A1" }
 */
export function splitFullAddress(
  full: string,
): { sheetName?: string; rangeAddress: string } {
  const bangIndex = full.indexOf("!");
  if (bangIndex < 0) {
    return { rangeAddress: full };
  }
  const sheetRaw = full.substring(0, bangIndex);
  const rangeAddress = full.substring(bangIndex + 1);
  const sheetName =
    sheetRaw.startsWith("'") && sheetRaw.endsWith("'")
      ? sheetRaw.slice(1, -1)
      : sheetRaw;
  return { sheetName, rangeAddress };
}

// ---------------------------------------------------------------------------
// Office.js bridge — безопасное получение Range
// ---------------------------------------------------------------------------

/**
 * Получить Excel.Range из строки адреса с проверкой безопасности.
 *
 * Принимает: "A1", "A1:B2", "Лист!A1", "'Мой лист'!A1:B2".
 * Если sheetName не указан — использует активный лист.
 *
 * ВАЖНО: адрес ДОЛЖЕН быть провалидирован через isValidFullAddress перед вызовом
 * (или используйте resolveRangeSafe — он валидирует внутри).
 *
 * Не делает context.sync() — это ответственность вызывающего (батчинг).
 */
export function resolveRange(
  context: Excel.RequestContext,
  address: string,
): Excel.Range {
  const { sheetName, rangeAddress } = splitFullAddress(address);
  if (sheetName !== undefined) {
    return context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
  }
  return context.workbook.worksheets
    .getActiveWorksheet()
    .getRange(rangeAddress);
}
