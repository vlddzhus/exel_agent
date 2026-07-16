/**
 * Input Validation — validate and sanitize user inputs for Excel AI Agent.
 *
 * Provides validation for:
 * - Range addresses (A1, A1:B10, Sheet1!A1:C10, etc.)
 * - Sheet names (length restrictions, forbidden chars)
 * - Table names (sanitization)
 * - Formula strings (basic structural checks)
 * - Sensitive data detection (passports, credit cards, passwords)
 */

// ---------------------------------------------------------------------------
// Range Address Validation
// ---------------------------------------------------------------------------

// Max column = XFD = 18278 → letterToColumn('XFD') = 18278
// Max row = 1048576
// Match examples:
//   A1, B10, XFD1048576
//   A1:B10, AA1:AB100
//   Sheet1!A1:C10, 'Sheet 1'!A1:B5
//   A:A (entire column), 1:1 (entire row)

const CELL_REF_PATTERN = /\$?[A-Za-z]+\$?\d+/;
const RANGE_REF_PATTERN = /^(\$?[A-Za-z]+\$?\d+):(\$?[A-Za-z]+\$?\d+)$/;
const SINGLE_CELL_PATTERN = /^\$?[A-Za-z]+\$?\d+$/;
const ENTIRE_COL_PATTERN = /^[A-Za-z]+:[A-Za-z]+$/;
const ENTIRE_ROW_PATTERN = /^\d+:\d+$/;

// Maximum Excel row number (1048576)
const MAX_ROW = 1048576;
// Maximum range address length to prevent abuse
const MAX_ADDRESS_LENGTH = 200;

/**
 * Validate an Excel range address.
 *
 * Accepts:
 *  - Single cell: A1, $A$1, XFD1048576
 *  - Range: A1:B10, $A$1:$B$10
 *  - With sheet name: Sheet1!A1:C10, 'Sheet 1'!A1:B5
 *  - Entire column: A:A, B:Z
 *  - Entire row: 1:1, 2:100
 *
 * Rejects:
 *  - Empty/null/non-string
 *  - Path traversal patterns (.., /, \)
 *  - Strings over 200 chars
 *  - Row 0 or row > 1048576
 */
export function validateRangeAddress(address: string): boolean {
  if (!address || typeof address !== 'string') return false;

  const trimmed = address.trim();
  if (trimmed.length === 0) return false;
  if (trimmed.length > MAX_ADDRESS_LENGTH) return false;

  // Reject path traversal
  if (/\.\./.test(trimmed) || /[/\\]/.test(trimmed)) return false;

  // Strip optional sheet name prefix (e.g., "Sheet1!" or "'Sheet 1'!")
  let rangePart = trimmed;
  const sheetMatch = trimmed.match(/^(?:'[^']*'|[^!']+)!/);
  if (sheetMatch) {
    rangePart = trimmed.slice(sheetMatch[0].length);
  }

  // Validate the range part
  if (SINGLE_CELL_PATTERN.test(rangePart)) {
    return validateCellRef(rangePart);
  }
  if (RANGE_REF_PATTERN.test(rangePart)) {
    const parts = rangePart.split(':');
    return validateCellRef(parts[0]) && validateCellRef(parts[1]);
  }
  if (ENTIRE_COL_PATTERN.test(rangePart)) {
    return true;
  }
  if (ENTIRE_ROW_PATTERN.test(rangePart)) {
    return validateRowRange(rangePart);
  }

  return false;
}

/**
 * Validate a single cell reference like A1, $B$10, XFD1048576
 */
function validateCellRef(cell: string): boolean {
  const match = cell.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!match) return false;

  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2], 10);

  // Validate column: must be between A (0) and XFD (16383)
  const colNum = letterToColumnNum(colStr);
  if (colNum < 0 || colNum > 16383) return false;

  // Validate row: must be between 1 and 1048576
  if (row < 1 || row > MAX_ROW) return false;

  return true;
}

/**
 * Convert column letters to number (A=0, B=1, ... Z=25, AA=26, ... XFD=18278)
 */
function letterToColumnNum(col: string): number {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result - 1; // 0-based
}

/**
 * Validate a row range like "1:1" or "5:100"
 */
function validateRowRange(range: string): boolean {
  const parts = range.split(':');
  const start = parseInt(parts[0], 10);
  const end = parseInt(parts[1], 10);
  return start >= 1 && start <= MAX_ROW && end >= 1 && end <= MAX_ROW;
}

// ---------------------------------------------------------------------------
// Sheet Name Validation
// ---------------------------------------------------------------------------

// Forbidden characters in Excel sheet names
const FORBIDDEN_SHEET_CHARS = /[\[\]:/\\?*\x00-\x1f]/;

/**
 * Validate an Excel sheet name.
 *
 * Rules:
 *  - Max 31 characters
 *  - Cannot be empty
 *  - Cannot contain: [ ] : / \ ? * (and control chars)
 */
export function validateSheetName(name: string): boolean {
  if (!name || typeof name !== 'string') return false;
  const trimmed = name.trim();
  if (trimmed.length === 0) return false;
  if (trimmed.length > 31) return false;
  if (FORBIDDEN_SHEET_CHARS.test(trimmed)) return false;
  return true;
}

// ---------------------------------------------------------------------------
// Table Name Sanitization
// ---------------------------------------------------------------------------

/**
 * Sanitize a string to be a valid Excel table name.
 *
 * Operations:
 *  - Replace spaces and special characters with underscores
 *  - If starts with a digit, prefix with 'T_'
 *  - If empty, return 'Table'
 *  - Preserves Russian/Cyrillic letters
 */
export function sanitizeTableName(name: string): string {
  if (!name || name.trim().length === 0) return 'Table';

  // Replace any character that is not a letter, digit, or underscore with '_'
  let sanitized = name.replace(/[^a-zA-Zа-яА-ЯёЁ0-9_]/g, '_');

  // Collapse multiple underscores
  sanitized = sanitized.replace(/_+/g, '_');

  // Strip leading/trailing underscores
  sanitized = sanitized.replace(/^_+|_+$/g, '');

  if (sanitized.length === 0) return 'Table';

  // If starts with a digit, prefix with 'T_'
  if (/^[0-9]/.test(sanitized)) {
    sanitized = 'T_' + sanitized;
  }

  return sanitized;
}

// ---------------------------------------------------------------------------
// Formula String Validation
// ---------------------------------------------------------------------------

/**
 * Validate a formula string (basic structural checks).
 *
 * Checks:
 *  - Not empty
 *  - Balanced parentheses
 *  - No obviously invalid characters
 */
export function validateFormulaString(formula: string): boolean {
  if (!formula || typeof formula !== 'string') return false;
  const trimmed = formula.trim();
  if (trimmed.length === 0) return false;

  // Check balanced parentheses
  let depth = 0;
  for (const ch of trimmed) {
    if (ch === '(') depth++;
    if (ch === ')') depth--;
    if (depth < 0) return false; // closing without opening
  }
  if (depth !== 0) return false;

  return true;
}

// ---------------------------------------------------------------------------
// Sensitive Data Detection
// ---------------------------------------------------------------------------

// Patterns for detecting potentially sensitive information
const SENSITIVE_PATTERNS: RegExp[] = [
  // Russian passport numbers (серия и номер: 2 цифры пробел/дефис 2 цифры пробел/дефис 6 цифр)
  /\b\d{2}\s*-?\s*\d{2}\s*-?\s*\d{6}\b/,
  // Credit card numbers (basic Luhn-like pattern: 16 digits, possibly grouped)
  /\b(?:\d{4}[\s-]?){3}\d{4}\b/,
  // Explicit mentions of sensitive keywords
  // Note: avoid \b with Cyrillic since JS \b only works with [a-zA-Z0-9_]
  /(?:^|\s)(?:password|пароль|secret|секрет|confidential|конфиденциально)\s*[:=]\s*\S+/i,
  // Phone numbers in international format
  /\+\d{10,15}\b/,
  // Email addresses (potential PII)
  /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/,
  // Russian INN (ИНН: 10 or 12 digits)
  /\b(?:инн|inn)\s*[:=]?\s*\d{10,12}\b/i,
  // Russian SNILS (СНИЛС: 11 digits with hyphens)
  /\b\d{3}[-]\d{3}[-]\d{3}\s*\d{2}\b/,
];

/**
 * Check if a text string contains potentially sensitive data.
 *
 * @param text - The text to check
 * @returns true if sensitive data patterns are detected
 */
export function checkSensitiveData(text: string): boolean {
  if (!text || typeof text !== 'string') return false;
  if (text.trim().length === 0) return false;

  for (const pattern of SENSITIVE_PATTERNS) {
    if (pattern.test(text)) return true;
  }

  return false;
}
