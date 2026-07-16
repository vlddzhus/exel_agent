/**
 * Unit tests for input-validation.ts — pure validation/sanitization functions.
 */
import {
  validateRangeAddress,
  validateSheetName,
  sanitizeTableName,
  validateFormulaString,
  checkSensitiveData,
} from '../src/taskpane/tools/input-validation';

// ===========================================================================
// validateRangeAddress
// ===========================================================================

describe('validateRangeAddress', () => {
  // ── Valid addresses ──

  test('valid: single cell A1', () => {
    expect(validateRangeAddress('A1')).toBe(true);
  });

  test('valid: range A1:B10', () => {
    expect(validateRangeAddress('A1:B10')).toBe(true);
  });

  test('valid: with sheet name', () => {
    expect(validateRangeAddress('Sheet1!A1:C10')).toBe(true);
  });

  test('valid: sheet with spaces', () => {
    expect(validateRangeAddress("'Sheet 1'!A1:B5")).toBe(true);
  });

  test('valid: XFD1048576 (max column + row)', () => {
    expect(validateRangeAddress('XFD1048576')).toBe(true);
  });

  test('valid: entire column A:A', () => {
    expect(validateRangeAddress('A:A')).toBe(true);
  });

  test('valid: entire row 1:1', () => {
    expect(validateRangeAddress('1:1')).toBe(true);
  });

  test('valid: multiple row range 1:100', () => {
    expect(validateRangeAddress('1:100')).toBe(true);
  });

  test('valid: entire column range A:Z', () => {
    expect(validateRangeAddress('A:Z')).toBe(true);
  });

  test('valid: absolute ref $A$1', () => {
    expect(validateRangeAddress('$A$1')).toBe(true);
  });

  test('valid: mixed absolute $A1', () => {
    expect(validateRangeAddress('$A1')).toBe(true);
  });

  test('valid: mixed absolute A$1', () => {
    expect(validateRangeAddress('A$1')).toBe(true);
  });

  test('valid: range with absolute refs', () => {
    expect(validateRangeAddress('$A$1:$B$10')).toBe(true);
  });

  test('valid: lowercase letters', () => {
    expect(validateRangeAddress('a1:b10')).toBe(true);
  });

  test('valid: Z1 (last single letter column)', () => {
    expect(validateRangeAddress('Z1')).toBe(true);
  });

  // ── Invalid addresses ──

  test('invalid: empty string', () => {
    expect(validateRangeAddress('')).toBe(false);
  });

  test('invalid: null', () => {
    expect(validateRangeAddress(null as any)).toBe(false);
  });

  test('invalid: undefined', () => {
    expect(validateRangeAddress(undefined as any)).toBe(false);
  });

  test('invalid: garbage string', () => {
    expect(validateRangeAddress('!@#$%^')).toBe(false);
  });

  test('invalid: path traversal with ..', () => {
    expect(validateRangeAddress('../../file')).toBe(false);
  });

  test('invalid: path traversal with slash', () => {
    expect(validateRangeAddress('Sheet1/../file!A1')).toBe(false);
  });

  test('invalid: backslash in address', () => {
    expect(validateRangeAddress('Sheet1\\A1')).toBe(false);
  });

  test('invalid: too long', () => {
    expect(validateRangeAddress('A'.repeat(201))).toBe(false);
  });

  test('invalid: row 0', () => {
    expect(validateRangeAddress('A0')).toBe(false);
  });

  test('invalid: row 1048577', () => {
    expect(validateRangeAddress('A1048577')).toBe(false);
  });

  test('invalid: column beyond XFD', () => {
    expect(validateRangeAddress('AAAA1')).toBe(false);
  });

  test('invalid: negative row', () => {
    expect(validateRangeAddress('A-1')).toBe(false);
  });

  test('invalid: letters only no numbers', () => {
    expect(validateRangeAddress('ABC')).toBe(false);
  });

  test('invalid: numbers only', () => {
    expect(validateRangeAddress('123')).toBe(false);
  });

  test('invalid: malformed sheet prefix', () => {
    expect(validateRangeAddress("'Sheet1!A1")).toBe(false);
  });
});

// ===========================================================================
// validateSheetName
// ===========================================================================

describe('validateSheetName', () => {
  // ── Valid names ──

  test('valid: normal sheet name', () => {
    expect(validateSheetName('Sheet1')).toBe(true);
  });

  test('valid: Russian name', () => {
    expect(validateSheetName('Лист1')).toBe(true);
  });

  test('valid: 31 characters (max allowed)', () => {
    expect(validateSheetName('A'.repeat(31))).toBe(true);
  });

  test('valid: name with spaces', () => {
    expect(validateSheetName('My Sheet')).toBe(true);
  });

  test('valid: name with numbers', () => {
    expect(validateSheetName('Data 2024')).toBe(true);
  });

  test('valid: name with underscores', () => {
    expect(validateSheetName('My_Sheet')).toBe(true);
  });

  test('valid: single character', () => {
    expect(validateSheetName('A')).toBe(true);
  });

  test('valid: name with dots', () => {
    expect(validateSheetName('Sheet.1')).toBe(true);
  });

  test('valid: name with parentheses', () => {
    expect(validateSheetName('Sheet(1)')).toBe(true);
  });

  // ── Invalid names ──

  test('invalid: empty string', () => {
    expect(validateSheetName('')).toBe(false);
  });

  test('invalid: null', () => {
    expect(validateSheetName(null as any)).toBe(false);
  });

  test('invalid: undefined', () => {
    expect(validateSheetName(undefined as any)).toBe(false);
  });

  test('invalid: 32 characters', () => {
    expect(validateSheetName('A'.repeat(32))).toBe(false);
  });

  test('invalid: brackets []', () => {
    expect(validateSheetName('Sheet[1]')).toBe(false);
  });

  test('invalid: colon :', () => {
    expect(validateSheetName('Sheet:1')).toBe(false);
  });

  test('invalid: forward slash /', () => {
    expect(validateSheetName('Sheet/1')).toBe(false);
  });

  test('invalid: backslash \\', () => {
    expect(validateSheetName('Sheet\\1')).toBe(false);
  });

  test('invalid: question mark ?', () => {
    expect(validateSheetName('Sheet?1')).toBe(false);
  });

  test('invalid: asterisk *', () => {
    expect(validateSheetName('Sheet*1')).toBe(false);
  });

  test('invalid: control character', () => {
    expect(validateSheetName('Sheet\x001')).toBe(false);
  });
});

// ===========================================================================
// sanitizeTableName
// ===========================================================================

describe('sanitizeTableName', () => {
  // ── Normal names ──

  test('normal name unchanged', () => {
    expect(sanitizeTableName('MyTable')).toBe('MyTable');
  });

  test('name with spaces -> underscores', () => {
    expect(sanitizeTableName('My Table')).toBe('My_Table');
  });

  test('name with special chars', () => {
    expect(sanitizeTableName('Table#1')).toBe('Table_1');
  });

  test('russian letters preserved', () => {
    expect(sanitizeTableName('Таблица1')).toBe('Таблица1');
  });

  test('name with multiple underscores collapsed', () => {
    expect(sanitizeTableName('My__Table')).toBe('My_Table');
  });

  test('name with leading/trailing underscores stripped', () => {
    expect(sanitizeTableName('_MyTable_')).toBe('MyTable');
  });

  // ── Edge cases ──

  test('empty -> Table', () => {
    expect(sanitizeTableName('')).toBe('Table');
  });

  test('null -> Table', () => {
    expect(sanitizeTableName(null as any)).toBe('Table');
  });

  test('undefined -> Table', () => {
    expect(sanitizeTableName(undefined as any)).toBe('Table');
  });

  test('starts with number -> prefix', () => {
    expect(sanitizeTableName('123Table')).toBe('T_123Table');
  });

  test('all special chars -> Table', () => {
    expect(sanitizeTableName('#$%^&')).toBe('Table');
  });

  test('whitespace only -> Table', () => {
    expect(sanitizeTableName('   ')).toBe('Table');
  });

  test('name with hyphens', () => {
    expect(sanitizeTableName('my-table')).toBe('my_table');
  });

  test('name with dots', () => {
    expect(sanitizeTableName('my.table')).toBe('my_table');
  });

  test('name with mixed case preserved', () => {
    expect(sanitizeTableName('MyTableName')).toBe('MyTableName');
  });

  test('name starting with underscore and number', () => {
    expect(sanitizeTableName('_1table')).toBe('T_1table');
  });
});

// ===========================================================================
// validateFormulaString
// ===========================================================================

describe('validateFormulaString', () => {
  test('valid simple formula', () => {
    expect(validateFormulaString('SUM(A1:A10)')).toBe(true);
  });

  test('valid nested formula', () => {
    expect(validateFormulaString('IF(A1>0;SUM(B1:B10);0)')).toBe(true);
  });

  test('valid formula with no parentheses', () => {
    expect(validateFormulaString('A1+B2')).toBe(true);
  });

  test('valid simple number', () => {
    expect(validateFormulaString('42')).toBe(true);
  });

  test('valid cell reference', () => {
    expect(validateFormulaString('A1')).toBe(true);
  });

  // ── Invalid ──

  test('invalid: empty string', () => {
    expect(validateFormulaString('')).toBe(false);
  });

  test('invalid: null', () => {
    expect(validateFormulaString(null as any)).toBe(false);
  });

  test('invalid: undefined', () => {
    expect(validateFormulaString(undefined as any)).toBe(false);
  });

  test('invalid: unmatched opening paren', () => {
    expect(validateFormulaString('SUM(A1:A10')).toBe(false);
  });

  test('invalid: unmatched closing paren', () => {
    expect(validateFormulaString('SUM(A1:A10))')).toBe(false);
  });

  test('invalid: closing before opening', () => {
    expect(validateFormulaString(')A1(')).toBe(false);
  });

  test('invalid: all whitespace', () => {
    expect(validateFormulaString('   ')).toBe(false);
  });

  test('invalid: deeply nested unbalanced', () => {
    expect(validateFormulaString('(((A1+B2)')).toBe(false);
  });
});

// ===========================================================================
// checkSensitiveData
// ===========================================================================

describe('checkSensitiveData', () => {
  // ── Detected patterns ──

  test('passport pattern detected', () => {
    expect(checkSensitiveData('паспорт 1234 567890')).toBe(true);
  });

  test('credit card pattern detected', () => {
    expect(checkSensitiveData('credit card 4111 1111 1111 1111')).toBe(true);
  });

  test('password field detected', () => {
    expect(checkSensitiveData('password: secret123')).toBe(true);
  });

  test('пароль (Russian password) detected', () => {
    expect(checkSensitiveData('пароль = mypass123')).toBe(true);
  });

  test('email address detected', () => {
    expect(checkSensitiveData('user@example.com')).toBe(true);
  });

  test('international phone detected', () => {
    expect(checkSensitiveData('+79161234567')).toBe(true);
  });

  test('Russian INN detected', () => {
    expect(checkSensitiveData('ИНН 1234567890')).toBe(true);
  });

  test('confidential keyword detected', () => {
    expect(checkSensitiveData('confidential: donotshare')).toBe(true);
  });

  test('secret keyword detected', () => {
    expect(checkSensitiveData('secret=abc123')).toBe(true);
  });

  // ── Not flagged ──

  test('normal text not flagged', () => {
    expect(checkSensitiveData('Hello, how are you?')).toBe(false);
  });

  test('empty string not flagged', () => {
    expect(checkSensitiveData('')).toBe(false);
  });

  test('null not flagged', () => {
    expect(checkSensitiveData(null as any)).toBe(false);
  });

  test('undefined not flagged', () => {
    expect(checkSensitiveData(undefined as any)).toBe(false);
  });

  test('simple numbers not flagged', () => {
    expect(checkSensitiveData('42')).toBe(false);
  });

  test('short number sequence not flagged', () => {
    expect(checkSensitiveData('12345')).toBe(false);
  });

  test('plain text without keywords not flagged', () => {
    expect(checkSensitiveData('The quick brown fox jumps over the lazy dog')).toBe(false);
  });

  test('table data not flagged', () => {
    expect(checkSensitiveData('Product, Price, Quantity')).toBe(false);
  });

  test('date not flagged', () => {
    expect(checkSensitiveData('2024-01-15')).toBe(false);
  });
});
