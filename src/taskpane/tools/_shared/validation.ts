/**
 * _shared/validation.ts — единая точка валидации аргументов инструментов.
 *
 * Все инструменты обязаны валидировать аргументы через этот модуль ДО передачи
 * в Office.js (см. docs/03-TOOLS-SPEC.md §3.2, docs/00-MASTER-PLAN.md P4).
 *
 * Что включено:
 *   - assertFullAddress / assertSheetName / sanitizeTableName (бросают)
 *   - detectSensitiveData / hasSensitiveData — детектор ПДн (паспорт/карта/ИНН/...)
 *   - validateAndSplitAddress — guard для write-инструментов
 *
 * Принципы:
 *   - Path traversal (../, /, \) недопустим — защита от записи вне диапазонов.
 *   - Имена листов/таблиц соответствуют ограничениям Excel.
 *   - ПДн детектируются перед отправкой в LLM-контекст (data protection).
 */

import {
  isValidFullAddress,
  isValidSheetName,
  splitFullAddress,
} from "./address";

// ---------------------------------------------------------------------------
// Address validation (бросают)
// ---------------------------------------------------------------------------

/**
 * Валидирует полный адрес (с опциональным именем листа) и бросает Error,
 * если адрес невалиден или содержит path traversal.
 *
 * Использовать ПЕРЕД каждым resolveRange / getRange.
 *
 * "A1" ✓  "A1:B2" ✓  "Лист!A1" ✓  "../../etc/passwd" ✗ (бросает)
 */
export function assertFullAddress(address: string): void {
  if (!isValidFullAddress(address)) {
    throw new Error(
      `Невалидный или небезопасный адрес диапазона: "${address}". ` +
        `Адрес должен быть вида "A1", "A1:B2" или "Лист!A1:B2".`,
    );
  }
}

/**
 * Валидирует имя листа. Бросает Error если невалидно.
 * Использовать перед worksheets.getItem / worksheets.add.
 */
export function assertSheetName(name: string): void {
  if (!isValidSheetName(name)) {
    throw new Error(
      `Невалидное имя листа: "${name}". ` +
        `Имя должно быть 1-31 символ без символов : \\ / ? * [ ].`,
    );
  }
}

/**
 * Санитизирует имя таблицы для Excel: обрезает, удаляет запрещённые символы,
 * ограничивает длину 31 символом. Возвращает безопасное имя или бросает Error.
 */
export function sanitizeTableName(name: string): string {
  if (!name || typeof name !== "string") {
    throw new Error("Имя таблицы обязательно");
  }
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Имя таблицы не может быть пустым");
  }
  const cleaned = trimmed.replace(/[\\/:?*[\]]/g, "").slice(0, 31);
  if (cleaned.length === 0) {
    throw new Error(
      `Имя таблицы после очистки стало пустым. Исходное: "${name}"`,
    );
  }
  return cleaned;
}

// ---------------------------------------------------------------------------
// Sensitive Data Detection (ПДн) — для защиты данных пользователя
// ---------------------------------------------------------------------------

/**
 * Регэкспы для распознавания российских ПДн.
 * Используется перед отправкой данных таблицы в LLM-контекст.
 *
 * ВАЖНО: это эвристика (false positives/negatives возможны). Цель — не идеальная
 * классификация, а редуцирование риска случайной утечки.
 */
const SENSITIVE_PATTERNS: ReadonlyArray<{ name: string; pattern: RegExp }> = [
  // Паспорт РФ: серии и номер в форматах "12 34 567890" или "1234 567890"
  { name: "passport", pattern: /\b\d{2}\s?\d{2}\s?\d{6}\b/ },
  // Номер банковской карты: 13-19 цифр, опционально с пробелами/дефисами
  { name: "card", pattern: /\b(?:\d[\s-]?){13,19}\b/ },
  // ИНН физлица (12 цифр) или юрлица (10 цифр)
  { name: "inn", pattern: /\b\d{10}|\d{12}\b/ },
  // СНИЛС: "123-456-789 00" или "12345678900"
  {
    name: "snils",
    pattern: /\b\d{3}-?\d{3}-?\d{3}[\s-]?\d{2}\b/,
  },
  // Российский телефон: +7/8 + 10 цифр в разных форматах.
  // Покрывает: +7 (495) 123-45-67, 89991234567, +7-999-123-45-67, 8 999 1234567
  // ВНИМАНИЕ: \b в начале НЕ работает (после +), поэтому опускаем — pattern
  // достаточно специфичный (\+7|8 + 10 цифр), чтобы не давать ложные срабатывания.
  {
    name: "phone_ru",
    pattern: /(?:\+7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}/,
  },
  // Email
  {
    name: "email",
    pattern: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/,
  },
];

export interface SensitiveDataFinding {
  type: string;
  match: string;
}

/**
 * Сканирует текст на наличие ПДн. Возвращает список находок.
 * Пустой массив = ПДн не обнаружено.
 */
export function detectSensitiveData(text: string): SensitiveDataFinding[] {
  if (!text || typeof text !== "string" || text.length === 0) return [];

  const findings: SensitiveDataFinding[] = [];
  for (const { name, pattern } of SENSITIVE_PATTERNS) {
    const match = pattern.exec(text);
    if (match) {
      findings.push({ type: name, match: match[0] });
    }
  }
  return findings;
}

/**
 * Быстрая проверка (boolean) — есть ли ПДн в тексте.
 * Удобно для guard-условий.
 */
export function hasSensitiveData(text: string): boolean {
  return detectSensitiveData(text).length > 0;
}

// ---------------------------------------------------------------------------
// Combined guard для инструментов
// ---------------------------------------------------------------------------

/**
 * Удобный guard для write-инструментов: валидирует адрес и возвращает
 * распарсенный { sheetName?, rangeAddress } для дальнейшего использования.
 *
 * Бросает Error при невалидном адресе.
 */
export function validateAndSplitAddress(
  fullAddress: string,
): { sheetName?: string; rangeAddress: string } {
  assertFullAddress(fullAddress);
  return splitFullAddress(fullAddress);
}
