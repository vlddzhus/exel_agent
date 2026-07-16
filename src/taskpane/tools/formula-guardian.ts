/**
 * Formula Guardian — validation, fixing, and utility functions for Excel formulas.
 *
 * Предоставляет валидацию формул, авто-исправление (вставка "*" между смежными
 * ссылками), проверку скобок, нормализацию регистра английских функций (сохраняя
 * русские имена функций).
 */

export interface ValidationResult {
  valid: boolean;
  fixedFormula?: string;
  error?: string;
}

// ---------------------------------------------------------------------------
// Column / Cell / Range Helpers — RE-EXPORT из единого источника истины.
//
// ВАЖНО: реальная реализация находится в ./_shared/address.ts (см.
// docs/03-TOOLS-SPEC.md §3.1, docs/00-MASTER-PLAN.md P6). Этот re-export
// сохраняет обратную совместимость с кодом, который уже импортирует адресные
// хелперы из formula-guardian. Новый код ДОЛЖЕН импортировать напрямую из
// ./_shared/address.
// ---------------------------------------------------------------------------

export {
  columnToLetter,
  letterToColumn,
  computeRange,
  parseCellAddress,
  parseRangeAddress,
} from "./_shared/address";

// ---------------------------------------------------------------------------
// Formula Validation & Fixing
// ---------------------------------------------------------------------------

/**
 * Множество русских имен функций Excel, которые НЕ нужно преобразовывать
 * в верхний регистр (они уже могут быть введены пользователем).
 * Английские имена (sum, if, vlookup) будут приведены к SUM, IF, VLOOKUP.
 */
const RUSSIAN_FUNCTIONS = new Set([
  'СУММ', 'СРЗНАЧ', 'СЧЁТ', 'СЧЁТЗ', 'МАКС', 'МИН', 'ОКРУГЛ', 'СУММПРОИЗВ',
  'СУММЕСЛИ', 'СУММЕСЛИМН', 'ВПР', 'ГПР', 'ИНДЕКС', 'ПОИСКПОЗ', 'ПРОСМОТРХ',
  'ЕСЛИ', 'И', 'ИЛИ', 'НЕ', 'ЕСЛИОШИБКА', 'ЕСНД', 'ЕСЛИМН', 'ПЕРЕКЛ',
  'СЦЕПИТЬ', 'ЛЕВСИМВ', 'ПРАВСИМВ', 'ПСТР', 'ДЛСТР', 'НАЙТИ', 'ЗАМЕНИТЬ',
  'ПОДСТАВИТЬ', 'ТЕКСТ', 'ОБЪЕДИНИТЬ', 'СЕГОДНЯ', 'ТДАТА', 'ДАТА', 'РАЗНДАТ',
  'ДЕНЬ', 'МЕСЯЦ', 'ГОД', 'СТАНДОТКЛОН', 'МЕДИАНА', 'КВАРТИЛЬ',
  'СТАНДОТКЛОНП', 'ДИСП', 'ABS', 'СТЕПЕНЬ', 'КОРЕНЬ', 'ОСТАТ', 'ЦЕЛОЕ',
  'ПРОИЗВЕД',
]);

/**
 * Validate and optionally fix a formula string.
 *
 * Steps:
 * 1. Uppercase English function names (sum→SUM), preserve Russian names (СУММ)
 * 2. Insert "*" between adjacent cell references (B8B6 → B8*B6)
 * 3. Check balanced parentheses
 *
 * @param formula — Formula WITHOUT leading "="
 * @returns ValidationResult with valid flag, optional fixed formula, or error
 */
export function validateFormula(formula: string): ValidationResult {
  if (!formula || formula.trim().length === 0) {
    return { valid: false, error: 'Formula is empty' };
  }

  let fixed = formula.trim();

  // 1. Uppercase English function names, preserve Russian
  fixed = uppercaseFunctionNames(fixed);

  // 2. Fix adjacent cell references without operator → insert "*"
  //    Pattern: две ссылки на ячейки подряд без оператора между ними
  //    Supports absolute references with $ (e.g., $A$1B2 → $A$1*B2)
  //    Uses lookahead to handle 3+ adjacent refs (A1B2C3 → A1*B2*C3)
  // Вставляем '*' между смежными cell-refs без оператора.
  // Lookahead (?=\$?[A-Z]+\$?\d+) требует, чтобы сразу после первого cell-ref
  // шёл второй cell-ref. Если между ними оператор (=, <>, +, -, *, / и т.д.) —
  // lookahead не матчит, '*' не вставляется (фикс бага A1=B2 → A1*=B2).
  fixed = fixed.replace(/(\$?[A-Z]+\$?\d+)(?=\$?[A-Z]+\$?\d+)/gi, '$1*');

  // 3. Check balanced parentheses
  if (!hasBalancedParentheses(fixed)) {
    return {
      valid: false,
      fixedFormula: fixed,
      error: `Несбалансированные скобки в формуле: "${formula}"`,
    };
  }

  return { valid: true, fixedFormula: fixed };
}

/**
 * Приводит английские имена функций к верхнему регистру,
 * но оставляет русские имена как есть (сохраняя регистр пользователя).
 */
function uppercaseFunctionNames(formula: string): string {
  return formula.replace(/([a-zA-Zа-яА-ЯёЁ]+)(?=\()/g, (match) => {
    const upper = match.toUpperCase();
    // Если это русское имя функции — не трогаем
    if (RUSSIAN_FUNCTIONS.has(upper) && /[а-яА-ЯёЁ]/.test(match)) {
      return match; // сохраняем как пользователь написал
    }
    // Иначе — приводим к верхнему регистру (английские функции)
    return upper;
  });
}

/**
 * Проверяет сбалансированность круглых скобок в выражении.
 */
function hasBalancedParentheses(formula: string): boolean {
  // ФИКС Фазы 1.6: игнорируем скобки внутри строковых литералов "...".
  // До фикса формула =")" считалась несбалансированной, а ="(" — нет.
  let depth = 0;
  let inString = false;
  for (const ch of formula) {
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue; // пропускаем символы внутри строки
    if (ch === '(') depth++;
    if (ch === ')') depth--;
    if (depth < 0) return false; // закрывающая без открывающей
  }
  return depth === 0;
}
