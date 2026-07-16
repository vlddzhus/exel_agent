/**
 * Unit tests for knowledge-tools.ts — embedded knowledge retrieval functions.
 *
 * These tests use only the embedded knowledge constants (no fetch required).
 * The fetch-dependent getKnowledge() function is NOT tested here since it
 * requires a browser/server environment.
 */
import { getEmbeddedKnowledge, searchKnowledge } from '../src/taskpane/tools/knowledge-tools';

// ===========================================================================
// getEmbeddedKnowledge
// ===========================================================================

describe('getEmbeddedKnowledge', () => {
  test('returns knowledge for excel-formulas category', async () => {
    const result = await getEmbeddedKnowledge('excel-formulas');
    expect(result).toBeTruthy();
    expect(result.length).toBeGreaterThan(0);
    expect(result).toContain('SUM');
  });

  test('returns knowledge for office-js-patterns category', async () => {
    const result = await getEmbeddedKnowledge('office-js-patterns');
    expect(result).toBeTruthy();
    expect(result.length).toBeGreaterThan(0);
    expect(result).toContain('getRange');
  });

  test('returns knowledge for agent-workflows category', async () => {
    const result = await getEmbeddedKnowledge('agent-workflows');
    expect(result).toBeTruthy();
    expect(result.length).toBeGreaterThan(0);
    expect(result).toContain('CRITICAL');
  });

  test('all category returns all knowledge combined', async () => {
    const result = await getEmbeddedKnowledge('all');
    expect(result).toBeTruthy();
    expect(result.length).toBeGreaterThan(500);
  });

  test('returns empty string for unknown category', async () => {
    const result = await getEmbeddedKnowledge('unknown-category');
    expect(result).toBe('');
  });

  test('excel-formulas contains IF function reference', async () => {
    const result = await getEmbeddedKnowledge('excel-formulas');
    expect(result).toContain('ЕСЛИ');
    expect(result).toContain('AVERAGE');
    expect(result).toContain('VLOOKUP');
    expect(result).toContain('Running Total');
  });

  test('office-js-patterns contains range operations', async () => {
    const result = await getEmbeddedKnowledge('office-js-patterns');
    expect(result).toContain('Range Operations');
    expect(result).toContain('Table Operations');
    expect(result).toContain('Formula Operations');
  });

  test('agent-workflows contains anti-patterns', async () => {
    const result = await getEmbeddedKnowledge('agent-workflows');
    expect(result).toContain('ANTI-PATTERNS');
    expect(result).toContain('NEVER');
    expect(result).toContain('Data Analysis');
  });
});

// ===========================================================================
// searchKnowledge
// ===========================================================================

describe('searchKnowledge', () => {
  test('finds SUM function', async () => {
    const results = await searchKnowledge('SUM');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('SUM'))).toBe(true);
  });

  test('finds СУММ (Russian)', async () => {
    const results = await searchKnowledge('СУММ');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('СУММ'))).toBe(true);
  });

  test('finds IF function', async () => {
    const results = await searchKnowledge('IF');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.toLowerCase().includes('if'))).toBe(true);
  });

  test('finds VLOOKUP', async () => {
    const results = await searchKnowledge('VLOOKUP');
    expect(results.length).toBeGreaterThan(0);
    // VLOOKUP content should contain "VLOOKUP"
    expect(results.some((r) => r.includes('VLOOKUP'))).toBe(true);
  });

  test('finds ВПР (Russian VLOOKUP)', async () => {
    const results = await searchKnowledge('ВПР');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('ВПР'))).toBe(true);
  });

  test('finds running total patterns', async () => {
    const results = await searchKnowledge('running total');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('Running Total'))).toBe(true);
  });

  test('finds percentage content', async () => {
    const results = await searchKnowledge('percent');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('Percentages'))).toBe(true);
  });

  test('finds XLOOKUP', async () => {
    const results = await searchKnowledge('XLOOKUP');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('XLOOKUP'))).toBe(true);
  });

  test('search is case-insensitive', async () => {
    const lowerResults = await searchKnowledge('vlookup');
    const upperResults = await searchKnowledge('VLOOKUP');
    expect(lowerResults.length).toBe(upperResults.length);
  });

  test('returns empty array for nonsense query', async () => {
    const results = await searchKnowledge('xyznonexistent12345');
    expect(results).toEqual([]);
  });

  test('returns empty array for empty query', async () => {
    const results = await searchKnowledge('');
    expect(results).toEqual([]);
  });

  test('returns empty array for whitespace query', async () => {
    const results = await searchKnowledge('   ');
    // After trim, this becomes '' which returns []
    expect(results).toEqual([]);
  });

  test('keyword match by partial word works', async () => {
    // "multiplication" is a keyword for the cell reference entry
    const results = await searchKnowledge('multiplication');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('CRITICAL'))).toBe(true);
  });

  test('finds date functions', async () => {
    const results = await searchKnowledge('TODAY');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('TODAY'))).toBe(true);
  });

  test('finds text functions', async () => {
    const results = await searchKnowledge('CONCATENATE');
    expect(results.length).toBeGreaterThan(0);
    expect(results.some((r) => r.includes('CONCATENATE'))).toBe(true);
  });
});
