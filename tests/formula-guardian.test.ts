/**
 * Unit tests for formula-guardian.ts — pure utility functions
 * (no Office.js dependency).
 */
import {
  validateFormula,
  columnToLetter,
  letterToColumn,
  computeRange,
  parseCellAddress,
  parseRangeAddress,
} from '../src/taskpane/tools/formula-guardian';

// ===========================================================================
// validateFormula
// ===========================================================================

describe('validateFormula', () => {
  // ── Adjacent cell references without operator ──

  test('fixes adjacent cell refs without operator', () => {
    const result = validateFormula('B8B6+B9B7');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('B8*B6+B9*B7');
  });

  test('fixes adjacent in complex formula', () => {
    const result = validateFormula('A1B2+C3D4');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1*B2+C3*D4');
  });

  test('passes through valid formula unchanged', () => {
    const result = validateFormula('B8*B6+B9*B7');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('B8*B6+B9*B7');
  });

  test('fixes multiple missing operators', () => {
    const result = validateFormula('A1B2C3');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1*B2*C3');
  });

  test('preserves existing operators when fixing', () => {
    const result = validateFormula('A1*B2C3+D4E5');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1*B2*C3+D4*E5');
  });

  test('handles dollar signs in cell refs (absolute refs)', () => {
    const result = validateFormula('$A$1B2');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('$A$1*B2');
  });

  // ── Edge cases ──

  test('handles empty formula', () => {
    const result = validateFormula('');
    expect(result.valid).toBe(false);
    expect(result.error).toContain('empty');
  });

  test('handles whitespace-only formula', () => {
    const result = validateFormula('   ');
    expect(result.valid).toBe(false);
    expect(result.error).toContain('empty');
  });

  test('handles formula with no cell refs', () => {
    const result = validateFormula('42');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('42');
  });

  test('handles single cell reference only', () => {
    const result = validateFormula('A1');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1');
  });

  test('handles formula with only operators and numbers', () => {
    const result = validateFormula('1+2*3/4');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('1+2*3/4');
  });

  test('handles formula with leading/trailing whitespace', () => {
    const result = validateFormula('  B8B6+B9B7  ');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('B8*B6+B9*B7');
  });

  // ── Function name casing ──

  test('uppercases English function names', () => {
    const result = validateFormula('sum(A1:A10)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('SUM(A1:A10)');
  });

  test('uppercases multiple English function names', () => {
    const result = validateFormula('if(A1>0,sum(B1:B10),0)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('IF(A1>0,SUM(B1:B10),0)');
  });

  test('preserves Russian function names', () => {
    const result = validateFormula('СУММ(A1:A10)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('СУММ(A1:A10)');
  });

  test('preserves lowercase Russian function names', () => {
    const result = validateFormula('сумм(A1:A10)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('сумм(A1:A10)');
  });

  test('handles mixed Russian and English functions', () => {
    const result = validateFormula('ЕСЛИ(A1>0;СУММ(B1:B10);0)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('ЕСЛИ(A1>0;СУММ(B1:B10);0)');
  });

  test('uppercases English functions in mixed context', () => {
    const result = validateFormula('if(A1>0;сумм(B1:B10);0)');
    expect(result.valid).toBe(true);
    // 'if' → 'IF' (English), 'сумм' → 'сумм' (Russian, preserved)
    expect(result.fixedFormula).toBe('IF(A1>0;сумм(B1:B10);0)');
  });

  test('handles nested English and Russian functions', () => {
    const result = validateFormula('ЕСЛИ(A1>0;СУММ(B1:B10);0)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('ЕСЛИ(A1>0;СУММ(B1:B10);0)');
  });

  // ── Parentheses balancing ──

  test('rejects unmatched opening parenthesis', () => {
    const result = validateFormula('SUM(A1:A10');
    expect(result.valid).toBe(false);
    expect(result.error).toContain('скобк');
  });

  test('rejects unmatched closing parenthesis', () => {
    const result = validateFormula('SUM(A1:A10))');
    expect(result.valid).toBe(false);
    expect(result.error).toContain('скобк');
  });

  test('handles nested parentheses', () => {
    const result = validateFormula('IF(A1>0;SUM(B1:B10);0)');
    expect(result.valid).toBe(true);
  });

  test('rejects deeply unbalanced parentheses', () => {
    const result = validateFormula('(((A1+B2)');
    expect(result.valid).toBe(false);
    expect(result.error).toContain('скобк');
  });

  test('rejects closing before opening', () => {
    const result = validateFormula('A1)+(B2');
    expect(result.valid).toBe(false);
    expect(result.error).toContain('скобк');
  });

  test('handles multiple nested parentheses', () => {
    const result = validateFormula('IF(A1>0;SUM(B1:B10);IF(C1<0;0;1))');
    expect(result.valid).toBe(true);
  });

  // ── Complex formulas ──

  test('handles column range in function argument', () => {
    const result = validateFormula('SUM(A1:INDEX(B:B;MATCH(C1;D:D;0)))');
    expect(result.valid).toBe(true);
  });

  test('handles formula with string literals', () => {
    const result = validateFormula('IF(A1="test";"hello";"world")');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('IF(A1="test";"hello";"world")');
  });

  test('handles formula with comparison operators', () => {
    const result = validateFormula('A1>=B2');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1>=B2');
  });

  test('handles formula with percentage', () => {
    const result = validateFormula('A1*10%');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1*10%');
  });

  // ============================================================
  // Фаза 1.6: регрессионные тесты (фиксы FormulaGuardian v2)
  // ============================================================

  test('Ф1.6: русская функция ЕСНД распознана (фикс опечатки ЕСЛИНД)', () => {
    const result = validateFormula('ЕСНД(A1;0)');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('ЕСНД(A1;0)');
  });

  test('Ф1.6: A1=B2 НЕ превращается в A1*=B2 (фикс оператора сравнения)', () => {
    const result = validateFormula('A1=B2');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1=B2');
  });

  test('Ф1.6: A1<>B2 НЕ превращается в A1*<>B2', () => {
    const result = validateFormula('A1<>B2');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('A1<>B2');
  });

  test('Ф1.6: IF(A1=B2; "да"; "нет") — оператор внутри функции сохраняется', () => {
    const result = validateFormula('IF(A1=B2;"да";"нет")');
    expect(result.valid).toBe(true);
    expect(result.fixedFormula).toBe('IF(A1=B2;"да";"нет")');
  });

  test('Ф1.6: скобки внутри строкового литерала не считаются', () => {
    // До фикса: ')' внутри строки ломал баланс скобок -> valid=false (баг)
    const result = validateFormula('IF(A1=1;"откр (";"закр )")');
    expect(result.valid).toBe(true);
  });

  test('Ф1.6: отдельная закрывающая скобка в строке =")" валидна', () => {
    const result = validateFormula('A1&")"');
    expect(result.valid).toBe(true);
  });
});

// ===========================================================================
// columnToLetter
// ===========================================================================

describe('columnToLetter', () => {
  test('0 returns A', () => {
    expect(columnToLetter(0)).toBe('A');
  });

  test('1 returns B', () => {
    expect(columnToLetter(1)).toBe('B');
  });

  test('25 returns Z', () => {
    expect(columnToLetter(25)).toBe('Z');
  });

  test('26 returns AA', () => {
    expect(columnToLetter(26)).toBe('AA');
  });

  test('51 returns AZ', () => {
    expect(columnToLetter(51)).toBe('AZ');
  });

  test('52 returns BA', () => {
    expect(columnToLetter(52)).toBe('BA');
  });

  test('701 returns ZZ', () => {
    expect(columnToLetter(701)).toBe('ZZ');
  });

  test('702 returns AAA', () => {
    expect(columnToLetter(702)).toBe('AAA');
  });

  test('16383 returns XFD (max Excel column)', () => {
    expect(columnToLetter(16383)).toBe('XFD');
  });
});

// ===========================================================================
// letterToColumn
// ===========================================================================

describe('letterToColumn', () => {
  test('A returns 0', () => {
    expect(letterToColumn('A')).toBe(0);
  });

  test('B returns 1', () => {
    expect(letterToColumn('B')).toBe(1);
  });

  test('Z returns 25', () => {
    expect(letterToColumn('Z')).toBe(25);
  });

  test('AA returns 26', () => {
    expect(letterToColumn('AA')).toBe(26);
  });

  test('AZ returns 51', () => {
    expect(letterToColumn('AZ')).toBe(51);
  });

  test('BA returns 52', () => {
    expect(letterToColumn('BA')).toBe(52);
  });

  test('ZZ returns 701', () => {
    expect(letterToColumn('ZZ')).toBe(701);
  });

  test('AAA returns 702', () => {
    expect(letterToColumn('AAA')).toBe(702);
  });

  test('XFD returns 16383', () => {
    expect(letterToColumn('XFD')).toBe(16383);
  });

  test('lowercase letters work', () => {
    expect(letterToColumn('aa')).toBe(26);
    expect(letterToColumn('xfd')).toBe(16383);
  });

  // ── Round-trip tests ──

  test('round-trip for all single letters', () => {
    for (let i = 0; i < 26; i++) {
      expect(letterToColumn(columnToLetter(i))).toBe(i);
    }
  });

  test('round-trip for double-letter columns', () => {
    const cols = [26, 50, 100, 500, 700, 701];
    for (const c of cols) {
      expect(letterToColumn(columnToLetter(c))).toBe(c);
    }
  });

  test('round-trip for triple-letter columns', () => {
    const cols = [702, 1000, 5000, 10000, 16383];
    for (const c of cols) {
      expect(letterToColumn(columnToLetter(c))).toBe(c);
    }
  });
});

// ===========================================================================
// computeRange
// ===========================================================================

describe('computeRange', () => {
  test('A11, 9 rows, 5 cols -> A11:E19', () => {
    expect(computeRange('A11', 9, 5)).toBe('A11:E19');
  });

  test('A1, 1 row, 1 col -> A1:A1', () => {
    expect(computeRange('A1', 1, 1)).toBe('A1:A1');
  });

  test('B2, 3 rows, 4 cols -> B2:E4', () => {
    expect(computeRange('B2', 3, 4)).toBe('B2:E4');
  });

  test('Z100, 1 row, 1 col -> Z100:Z100', () => {
    expect(computeRange('Z100', 1, 1)).toBe('Z100:Z100');
  });

  test('AA1, 2 rows, 2 cols -> AA1:AB2', () => {
    expect(computeRange('AA1', 2, 2)).toBe('AA1:AB2');
  });

  test('XFD1048576, 1 row, 1 col -> XFD1048576:XFD1048576', () => {
    expect(computeRange('XFD1048576', 1, 1)).toBe('XFD1048576:XFD1048576');
  });

  test('A1, 10 rows, 1 col -> A1:A10', () => {
    expect(computeRange('A1', 10, 1)).toBe('A1:A10');
  });

  test('A1, 1 row, 10 cols -> A1:J1', () => {
    expect(computeRange('A1', 1, 10)).toBe('A1:J1');
  });
});

// ===========================================================================
// parseCellAddress
// ===========================================================================

describe('parseCellAddress', () => {
  test('A1 -> {col:A, row:1}', () => {
    expect(parseCellAddress('A1')).toEqual({ col: 'A', row: 1 });
  });

  test('B10 -> {col:B, row:10}', () => {
    expect(parseCellAddress('B10')).toEqual({ col: 'B', row: 10 });
  });

  test('AA100 -> {col:AA, row:100}', () => {
    expect(parseCellAddress('AA100')).toEqual({ col: 'AA', row: 100 });
  });

  test('XFD1048576 -> {col:XFD, row:1048576}', () => {
    expect(parseCellAddress('XFD1048576')).toEqual({ col: 'XFD', row: 1048576 });
  });

  test('throws on empty string', () => {
    expect(() => parseCellAddress('')).toThrow();
  });

  test('throws on letters only', () => {
    expect(() => parseCellAddress('ABC')).toThrow();
  });

  test('throws on numbers only', () => {
    expect(() => parseCellAddress('123')).toThrow();
  });

  test('throws on null/undefined', () => {
    expect(() => parseCellAddress(null as any)).toThrow();
    expect(() => parseCellAddress(undefined as any)).toThrow();
  });

  test('throws on mixed invalid format', () => {
    expect(() => parseCellAddress('1A')).toThrow();
    expect(() => parseCellAddress('A1B2')).toThrow();
  });
});

// ===========================================================================
// parseRangeAddress
// ===========================================================================

describe('parseRangeAddress', () => {
  test('A1:C10 -> parsed correctly', () => {
    const result = parseRangeAddress('A1:C10');
    expect(result).toEqual({
      startCol: 'A', startRow: 1,
      endCol: 'C', endRow: 10,
    });
  });

  test('B2:E4 -> parsed correctly', () => {
    const result = parseRangeAddress('B2:E4');
    expect(result).toEqual({
      startCol: 'B', startRow: 2,
      endCol: 'E', endRow: 4,
    });
  });

  test('AA1:AB100 -> parsed correctly', () => {
    const result = parseRangeAddress('AA1:AB100');
    expect(result).toEqual({
      startCol: 'AA', startRow: 1,
      endCol: 'AB', endRow: 100,
    });
  });

  test('throws on single cell', () => {
    expect(() => parseRangeAddress('A1')).toThrow();
  });

  test('throws on empty string', () => {
    expect(() => parseRangeAddress('')).toThrow();
  });

  test('throws on malformed range', () => {
    expect(() => parseRangeAddress('A1:')).toThrow();
    expect(() => parseRangeAddress(':B2')).toThrow();
    expect(() => parseRangeAddress('A1:B2:C3')).toThrow();
  });
});
